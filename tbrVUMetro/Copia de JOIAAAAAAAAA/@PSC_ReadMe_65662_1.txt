Title: WaveIn Recorder
Description: Uses the waveIn API to record audio signals.
You can change the recording source while recording and change its volume. The recorded samples get saved to a WAV file, with full ACM support, so you can even use an MP3 codec, if installed.
The recorded samples get visualized through a frequency spectrum (Ulli's FFT class) and an amplitudes curve. Tested with XP. /// Update: Replaced the common dialog control with a class by Bill Bither (found at PSC), added some comments to the code, cleaned up a bit /// Update: fake window will be ceated only on start of recording, secured some parts in WaveInRecorder where pointers are used /// Update...: replaced subclassing with Paul Caton's technique /// update...: mixer functions should work fine now, uses a real mixer handle now ;), also the inverted line selection bug is fixed /// and once more...: mixer functions should be really stable now, added effects (echo, amplify, phase shift, 7 band graphical equalizer). Tipp: add echo, give it a short length (about 4-5 pixels to the left from the default) and set the echo amp to 6, makes your recording sound like a robot :)
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=65662&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
