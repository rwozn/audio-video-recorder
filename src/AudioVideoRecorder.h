#pragma once

#include <string>
#include <vector>
#include <dshow.h>
#include <atlbase.h> // CComPtr

// Exports class identifiers (CLSIDs) and interface identifiers (IIDs)
#pragma comment(lib, "Strmiids")

class AudioVideoRecorder
{
	const std::wstring outputFileName;

	// Values for the start and stop events, arbitrary values
	static const WORD RECORDING_START_COOKIE = 0xDEAD;
	static const WORD RECORDING_STOP_COOKIE = 0xBEEF;

	// returns EnumMoniker used for enumerating devices
	static HRESULT getEnumMoniker(CComPtr<IEnumMoniker>& enumMoniker, const IID& deviceCategory);

	// returns capture filter (either Video capture filter or Audio capture filter, depending on deviceCategory)
	static HRESULT getCaptureFilter(CComPtr<IBaseFilter>& captureFilter, const IID& deviceCategory);

	// returns GraphBuilder and CaptureGraphBuilder2 after preparation
	HRESULT setupCaptureGraph(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, DWORD recordingDuration);

	// adds Video capture filter to the graph builder; returns true if there's a video recording device (a camera), false otherwise
	static bool setupVideoCaptureFilter(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, CComPtr<IBaseFilter>& muxFilter, DWORD recordingDuration);

	// adds Audio capture filter to the graph builder; returns true if there's an audio recording device (a microphone), false otherwise
	static bool setupAudioCaptureFilter(CComPtr<IGraphBuilder>& graphBuilder, CComPtr<ICaptureGraphBuilder2>& captureGraphBuilder2, CComPtr<IBaseFilter>& muxFilter, DWORD recordingDuration);

public:
	AudioVideoRecorder(const std::wstring& outputFileName);

	~AudioVideoRecorder();

	// records for `duration` seconds
	void record(DWORD duration);

	std::wstring getOutputFileName();
};
