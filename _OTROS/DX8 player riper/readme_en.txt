
Monoton Sound Library

>> how to get it to run

1. compile monoton_ds_en.vbp.
2. needed files for the encoders/decoders:

WAV: -
MP3: MP3 ACM Codec (e.g. Lame), lame_enc.dll
WMA: WMF SDK/Runtime or WMP 9/10
OGG: ogg.dll, vorbis.dll
APE: MACDll.dll
CDA: (ASPI)

Notes on CDA:

Monoton supports digital playback of CD audio.
Therefore it needs low level access to the disc.
Win 9x/Me: Adaptec ASPI driver needed
Win NT/2k/XP: SPTI (needs admin rights), IOCTLs or ASPI

SPTI is a control code for DeviceIoControl(), which
allows full access to drives, but needs administrator priviledges.
CDROM IOCTLs for DeviceIoControl() do not need admin privs,
but are less functional then the SPTI.
Monoton may not select the ASPI on NT/2K/XP.

Where to get all these files:

APE and MP3:
http://actorics.de/rm_code/code/monoton/apempeg.zip
Ogg Vorbis:
http://actorics.de/rm_code/code/monoton/libov.zip
WMF Runtime:
http://actorics.de/rm_code/code/monoton/wmfdist.zip

3. run the examples.

Examples:

1. player: playback of the supported file types
2. ripper: rips audio CDs to wav/mp3/wma/ogg/ape
3. simple: simplest possible playback of files with monoton
4. transcoder: convert from one format to the other