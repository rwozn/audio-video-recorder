#include "AudioVideoRecorder.h"

#include <iostream>

int main()
{
	const DWORD duration = 10;

	AudioVideoRecorder recorder(L"record.avi");

	std::cout << "Recording " << duration << " seconds of audio and video (camera + microphone) to \"";

	std::wcout << recorder.getOutputFileName();

	std::cout << "\"..." << std::endl;

	recorder.record(duration);

	return 0;
}