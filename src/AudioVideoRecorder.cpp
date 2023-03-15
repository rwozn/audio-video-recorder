#include "AudioVideoRecorder.h"

#include <iostream>
#include <stdexcept>

HRESULT AudioVideoRecorder::getEnumMoniker(CComPtr<IEnumMoniker>& enumMoniker, const IID& deviceCategory)
{
	CComPtr<ICreateDevEnum> createDevEnum;

	// create the system device enumerator	
	HRESULT result = createDevEnum.CoCreateInstance(CLSID_SystemDeviceEnum, NULL, CLSCTX_INPROC_SERVER);

	if(FAILED(result))
		return result;

	// create an enumerator for the video input device category
	result = createDevEnum->CreateClassEnumerator(deviceCategory, &enumMoniker, 0);

	// the category is empty, treat as an error
	if(result == S_FALSE)
		result = VFW_E_NOT_FOUND;

	return result;
}

// possible values for deviceCategory:
// CLSID_AudioInputDeviceCategory - audio capture devices, e.g. microphone
// CLSID_VideoInputDeviceCategory - video capture devices, e.g. camera
// this functions returns the first device for the given category
HRESULT AudioVideoRecorder::getCaptureFilter(CComPtr<IBaseFilter>& captureFilter, const IID& deviceCategory)
{
	CComPtr<IEnumMoniker> enumMoniker;

	HRESULT result = getEnumMoniker(enumMoniker, deviceCategory);

	if(FAILED(result))
		return result;

	while(1)
	{
		CComPtr<IMoniker> moniker;

		if(enumMoniker->Next(1, &moniker, NULL) == S_FALSE)
			break;

		// get the first device
		result = moniker->BindToObject(NULL, NULL, IID_IBaseFilter, (void**)&captureFilter);

		if(SUCCEEDED(result))
			return result;
	}

	return E_FAIL;
}

HRESULT AudioVideoRecorder::setupCaptureGraph(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, DWORD recordingDuration)
{
	// create the capture graph builder
	HRESULT result = captureGraphBuilder2.CoCreateInstance(CLSID_CaptureGraphBuilder2, NULL, CLSCTX_INPROC_SERVER);

	if(FAILED(result))
		return result;

	// create the filter graph menager
	result = graphBuilder.CoCreateInstance(CLSID_FilterGraph, 0, CLSCTX_INPROC_SERVER);

	if(FAILED(result))
		return result;

	// initialize the capture graph builder
	result = captureGraphBuilder2->SetFiltergraph(graphBuilder);

	if(FAILED(result))
		return result;

	// Capture Filter => AVI Mux => File Writer

	// The AVI Mux filter takes the video stream from the capture pin and packages it into an AVI stream.
	// The File Writer filter writes the AVI stream to disk.
	// The AVI Mux filter accepts multiple input streams and interleaves them into AVI format.
	// The filter uses separate input pins for each input stream, and one output pin for the AVI stream.
	CComPtr<IBaseFilter> muxFilter;

	// This function cocreates the AVI Mux filter and the File Writer filter and adds them to the graph.
	// It also sets the file name on the File Writer filter and connects the two filters
	result = captureGraphBuilder2->SetOutputFileName(&MEDIASUBTYPE_Avi, outputFileName.c_str(), &muxFilter, NULL);

	if(FAILED(result))
		return result;

	bool cameraAvailable = setupVideoCaptureFilter(graphBuilder, captureGraphBuilder2, muxFilter, recordingDuration);

	bool microphoneAvailable = setupAudioCaptureFilter(graphBuilder, captureGraphBuilder2, muxFilter, recordingDuration);

	if(!cameraAvailable && !microphoneAvailable)
		throw std::runtime_error("No camera and microphone available");
	
	if(!cameraAvailable)
		std::cout << "No camera available" << std::endl;
	else if(!microphoneAvailable)
		std::cout << "No microphone available" << std::endl;

	// If you are capturing audio and video from two separate devices, it is a good idea to make the audio stream the master stream.
	// This helps to prevent drift between the two streams, because the AVI Mux filter adjust the playback rate on the video stream to match the audio stream.
	// To set the master stream, call the IConfigAviMux::SetMasterStream method on the AVI Mux filter
	if(cameraAvailable && microphoneAvailable)
	{
		CComPtr<IConfigAviMux> configAviMux;

		result = muxFilter->QueryInterface(IID_IConfigAviMux, (void**)&configAviMux);

		if(FAILED(result))
			return result;

		// The parameter to SetMasterStream is the stream number, which is determined by the order in 
		// which you call RenderStream. For example, if you call RenderStream first for video and
		// then for audio, the video is stream 0 and the audio is stream 1.
		result = configAviMux->SetMasterStream(1);

		if(FAILED(result))
			return result;

		// You may also want to set how the AVI Mux filter interleaves the audio and video streams,
		// by calling the IConfigInterleaving::put_Mode method:
		CComPtr<IConfigInterleaving> configInterleaving;

		result = muxFilter->QueryInterface(IID_IConfigAviMux, (void**)&configInterleaving);

		if(FAILED(result))
			return result;

		// INTERLEAVE_CAPTURE - the AVI Mux performs interleaving at a rate that is suitable for video capture.
		// INTERLEAVE_NONE - no interleaving — the AVI Mux will simply write the data in the order that it arrives.
		// INTERLEAVE_FULL - the AVI Mux performs full interleaving; however,
		// this mode is less suitable for video capture because it requires the most overheard.
		result = configInterleaving->put_Mode(INTERLEAVE_CAPTURE);

		if(FAILED(result))
			return result;
	}

	return S_OK;
}

bool AudioVideoRecorder::setupVideoCaptureFilter(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, CComPtr<IBaseFilter>& muxFilter, DWORD recordingDuration)
{
	CComPtr<IBaseFilter> videoCaptureFilter;

	HRESULT result = getCaptureFilter(videoCaptureFilter, CLSID_VideoInputDeviceCategory);

	// no camera
	if(FAILED(result))
		return false;

	result = graphBuilder->AddFilter(videoCaptureFilter, L"Capture Filter");

	if(FAILED(result))
		throw std::runtime_error("AddFilter failed for camera: " + std::to_string(result));

	// call the ICaptureGraphBuilder2::RenderStream method to connect the capture filter to the AVI Mux
	result = captureGraphBuilder2->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, videoCaptureFilter, NULL, muxFilter);

	if(FAILED(result))
	{
		graphBuilder->RemoveFilter(videoCaptureFilter);

		throw std::runtime_error("RenderStream failed for camera: " + std::to_string(result));
	}

	// define the times when the stream will start and stop,
	// relative to the time when the graph starts running
	REFERENCE_TIME recordingStartTime = 0;
	REFERENCE_TIME recordingStopTime = recordingDuration * 10000000; // multiply for correct units

	// Call IMediaControl::Run to run the graph
	// Until you run the graph, the ControlStream method has no effect.
	// The last two parameters are used for getting event notifications when the stream starts and stops.
	// For each stream that you control using this method, the filter graph sends a pair of events:
	// EC_STREAM_CONTROL_STARTED when the stream starts, and EC_STREAM_CONTROL_STOPPED when the stream stops.
	// The values of wStartCookie and wStopCookie are used as the second event parameter.
	// Thus, lParam2 in the start event equals wStartCookie, and lParam2 in the stop event equals wStopCookie.
	result = captureGraphBuilder2->ControlStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, videoCaptureFilter, &recordingStartTime, &recordingStopTime, RECORDING_START_COOKIE, RECORDING_STOP_COOKIE);

	if(FAILED(result))
	{
		graphBuilder->RemoveFilter(videoCaptureFilter);

		throw std::runtime_error("ControlStream failed for camera: " + std::to_string(result));
	}

	return true;
}

bool AudioVideoRecorder::setupAudioCaptureFilter(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, CComPtr<IBaseFilter>& muxFilter, DWORD recordingDuration)
{
	CComPtr<IBaseFilter> audioCaptureFilter;

	HRESULT result = getCaptureFilter(audioCaptureFilter, CLSID_AudioInputDeviceCategory);

	// no microphone
	if(FAILED(result))
		return false;

	result = graphBuilder->AddFilter(audioCaptureFilter, L"Capture Filter");

	if(FAILED(result))
		throw std::runtime_error("AddFilter failed for microphone: " + std::to_string(result));

	// call the ICaptureGraphBuilder2::RenderStream method to connect the capture filter to the AVI Mux
	result = captureGraphBuilder2->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, audioCaptureFilter, NULL, muxFilter);

	if(FAILED(result))
	{
		graphBuilder->RemoveFilter(audioCaptureFilter);

		throw std::runtime_error("RenderStream failed for microphone: " + std::to_string(result));
	}

	// define the times when the stream will start and stop,
	// relative to the time when the graph starts running
	REFERENCE_TIME recordingStartTime = 0;
	REFERENCE_TIME recordingStopTime = recordingDuration * 10000000; // multiply for correct units

	// Call IMediaControl::Run to run the graph					  
	// Until you run the graph, the ControlStream method has no effect.
	// The last two parameters are used for getting event notifications when the stream starts and stops.
	// For each stream that you control using this method, the filter graph sends a pair of events:
	// EC_STREAM_CONTROL_STARTED when the stream starts, and EC_STREAM_CONTROL_STOPPED when the stream stops.
	// The values of wStartCookie and wStopCookie are used as the second event parameter.
	// Thus, lParam2 in the start event equals wStartCookie, and lParam2 in the stop event equals wStopCookie.
	result = captureGraphBuilder2->ControlStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, audioCaptureFilter, &recordingStartTime, &recordingStopTime, RECORDING_START_COOKIE, RECORDING_STOP_COOKIE);

	if(FAILED(result))
	{
		graphBuilder->RemoveFilter(audioCaptureFilter);

		throw std::runtime_error("ControlStream failed for microphone: " + std::to_string(result));
	}

	return true;
}

AudioVideoRecorder::AudioVideoRecorder(const std::wstring& outputFileName):
	outputFileName(outputFileName)
{
	CoInitialize(NULL);
}

AudioVideoRecorder::~AudioVideoRecorder()
{
	// Unloads any DLLs that are no longer in use, probably because the DLL no longer has any instantiated COM objects outstanding.
	CoFreeUnusedLibraries();
	
	CoUninitialize();
}

void AudioVideoRecorder::record(DWORD duration)
{
	HRESULT result;

	CComPtr<IGraphBuilder> graphBuilder;

	CComPtr<ICaptureGraphBuilder2> captureGraphBuilder2;

	result = setupCaptureGraph(graphBuilder, captureGraphBuilder2, duration);
	
	if(FAILED(result))
		throw std::runtime_error("setupCaptureGraph failed with code " + std::to_string(result));

	CComPtr<IMediaEvent> mediaEvent;

	result = graphBuilder->QueryInterface(&mediaEvent);

	if(FAILED(result))
		throw std::runtime_error("QueryInterface failed for IMediaEvent with code " + std::to_string(result));

	CComPtr<IMediaControl> mediaControl;

	result = graphBuilder->QueryInterface(&mediaControl);

	if(FAILED(result))
		throw std::runtime_error("QueryInterface failed for IMediaControl with code " + std::to_string(result));

	result = mediaControl->Run();

	if(FAILED(result))
		throw std::runtime_error("IMediaControl->Run failed with code " + std::to_string(result));

	// This method blocks until there is an event to return or until a specified time elapses.
	// Assuming there is a queued event, the method returns with the event code and the two event parameters.
	// After calling GetEvent, an application should always call the IMediaEvent::FreeEventParams method to
	// release any resources associated with the event parameters. For example, a parameter might
	// be a BSTR value that was allocated by the filter graph.

	while(1)
	{
		long eventCode;

		LONG_PTR parameter1;
		LONG_PTR parameter2;

		result = mediaEvent->GetEvent(&eventCode, &parameter1, &parameter2, 0);

		// GetEvent returns either S_OK or E_ABORT
		// E_ABORT means there are no events in the queue
		bool finishedRecording = SUCCEEDED(result) && parameter2 == RECORDING_STOP_COOKIE;
		
		mediaEvent->FreeEventParams(eventCode, parameter1, parameter2);

		if(finishedRecording)
			break;
	}

	mediaControl->Stop();
}

std::wstring AudioVideoRecorder::getOutputFileName()
{
	return outputFileName;
}