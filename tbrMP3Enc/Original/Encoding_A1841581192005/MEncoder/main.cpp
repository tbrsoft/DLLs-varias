
#include <windows.h>
#include <stdio.h>
#include <io.h>
#include <fcntl.h>
#include <sys/stat.h>
#include "main.h"
#include "bladedll.h"


// MP3 encoding prototypes
BEINITSTREAM	beInitStream;
BEENCODECHUNK	beEncodeChunk;
BEDEINITSTREAM	beDeinitStream;
BECLOSESTREAM	beCloseStream;
BEVERSION		beVersion;
BE_CONFIG		b;
BE_ERR			err;


//Variables
int				encoding_still;
double			percent_done = 0;
DWORD			dwFileSize = 0;
DWORD			dwMP3Buffer = 0;
DWORD			dwMP3BufferSize = 0;
LONG            thisBitrate;
BOOL            thiscopyright = FALSE;
BOOL            thisoriginal = FALSE;
BOOL            thisprivate = FALSE;
BOOL            thiscrc = FALSE;
BOOL            cancel = FALSE;
long			myfreq = 0;
int				mybits	= 0;
int				mychannels = 0;
HANDLE			hFile;	


//The Lame DLL Instance
HINSTANCE hLameDLL;


// enum to MP3 encoding status makes this multithreaded :) used for progress
typedef BOOL (CALLBACK* ENUMENC) (int);

//Call WINEXPORT so VB Understands The Dll :0
#define WINEXPORT __declspec(dllexport) WINAPI


//All DLL's Have To Have DLLMain for Entry Point
extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID /*lpReserved*/)
{
	//If The DLL IS First Being Loaded
    if (dwReason == DLL_PROCESS_ATTACH)
    {
		//Load The Lame Encoder DLL ON Starting Point
		hLameDLL  = LoadLibrary("LAME_ENC.DLL");     
    }
	//If The DLL Is Being UnLoaded
    else if (dwReason == DLL_PROCESS_DETACH)
		
		// cleanup processing Unload Lame Encoder DLL
		if (hLameDLL)
			FreeLibrary(hLameDLL);
        
    return TRUE;    // ok
}


//SetBitrate
LONG WINEXPORT SetBitrate(LONG bit)
{
	thisBitrate = bit;
	return 0;

}


//Set Copyright Info
LONG WINEXPORT SetCopyright(BOOL cpy)
{
	thiscopyright = cpy;
	return 0;

}


//Set Original Info
LONG WINEXPORT SetOriginal(BOOL org)
{
	thisoriginal = org;
	return 0;

}


//Set Private Info
LONG WINEXPORT SetPrivate(BOOL priv)
{
	thisprivate = priv;
	return 0;

}



//Set CRC Info
LONG WINEXPORT SetCRC(BOOL crc)
{
	thiscrc = crc;
	return 0;

}


//Set Cancel
LONG WINEXPORT Cancel(BOOL cncl)
{
	cancel = cncl;
	return 0;
}


LONG GetInWaveProp(LPCTSTR infile)
{
	HMMIO hmmio;
	MMCKINFO mmckinfoParent;
	MMCKINFO mmckinfoSubchunk;
	WAVEFORMATEX Format;

	float dAudioLength = 0;

	hFile = CreateFile(infile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

	if (hFile == INVALID_HANDLE_VALUE)
		return 0;

	dwFileSize = GetFileSize(hFile, NULL);

	CloseHandle(hFile);


	// Open a handle to the file to be examined.
	if (!(hmmio = mmioOpen((LPSTR)infile, NULL, MMIO_READ | MMIO_ALLOCBUF)))
		return 0;

	// First, we check to see if it is a WAV file.
	mmckinfoParent.fccType = mmioStringToFOURCC("WAVE", 0);
	if (mmioDescend(hmmio, (LPMMCKINFO) &mmckinfoParent, NULL, MMIO_FINDRIFF))
	{
		mmioClose(hmmio, 0);
		return 0;
	}

	// Get format info
	mmckinfoSubchunk.ckid = mmioStringToFOURCC("fmt", 0);
	MMRESULT mmResult = mmioDescend(hmmio, &mmckinfoSubchunk, &mmckinfoParent, MMIO_FINDCHUNK);
	if (mmResult)
	{
		mmioClose(hmmio, 0);
		return 0;
	}

	LONG lRet = mmioRead(hmmio, (HPSTR) &Format, mmckinfoSubchunk.cksize);
	if (lRet == -1)
	{
		mmioClose(hmmio, 0);
		return 0;
	}

	// Find the data subchunk
	mmckinfoSubchunk.ckid = mmioStringToFOURCC("data", 0);
	mmResult = mmioDescend(hmmio, &mmckinfoSubchunk, &mmckinfoParent, MMIO_FINDCHUNK);
	if (mmResult)
	{
		mmioClose(hmmio, 0);
		return 0;
	}
	float dTemp = float(Format.wBitsPerSample / 8);
	dAudioLength = mmckinfoSubchunk.cksize / (Format.nSamplesPerSec * Format.nChannels * dTemp);

	myfreq = Format.nSamplesPerSec;
	mychannels = Format.nChannels;
	mybits = Format.wBitsPerSample;


	return 0;
}



LONG WINEXPORT EncodeMp3(LPCSTR lpszWavFile, ENUMENC &EnumEncoding)
{
	//Load The Lame Encoder DLL 
	hLameDLL  = LoadLibrary("LAME_ENC.DLL"); 

	//If There is no Wave File Passed to this dll then return a -1
	if(lpszWavFile == "")
	{
		return -1;
	}


	// If The Lame DLL Didn't Load correctly then return a -1
	if(!hLameDLL)
	{
		return -1;
	}


	// Get Lame Interface
	beInitStream = (BEINITSTREAM)GetProcAddress( HINSTANCE(hLameDLL), "beInitStream" );
	beEncodeChunk = (BEENCODECHUNK)GetProcAddress( HINSTANCE(hLameDLL), "beEncodeChunk" );
	beDeinitStream = (BEDEINITSTREAM)GetProcAddress( HINSTANCE(hLameDLL), "beDeinitStream" );
	beCloseStream = (BECLOSESTREAM)GetProcAddress( HINSTANCE(hLameDLL), "beCloseStream" );
	beVersion = (BEVERSION)GetProcAddress( HINSTANCE(hLameDLL), "beVersion" );

	//If any Part of the Lame Interface Didn't Load correctly then return a -1
	if(!beInitStream || !beEncodeChunk || !beDeinitStream || !beCloseStream)
	{
		return -1;
	}

	//Open The Wve file as ReadOnly In Binary Format
	int hIn = open(lpszWavFile, O_RDONLY | O_BINARY);

	//If The Wave File Didn't Load correctly then return a -1
	if(hIn == -1)
	{
		return -1;
	}
	
	//set the Output file (zOutoutFilename) to the max path of the file
	char zOutputFilename[MAX_PATH + 1];
	
	//copy string passed from the function call lpszWavFile to the char
	//zOutputFilename with its full max file path
	lstrcpy(zOutputFilename, lpszWavFile);

	//get the char zOutputFilename string length
	int l = lstrlen(zOutputFilename);

	//now start deleting the ending of the file name til it gets
	//to a period(.) and clip the . but go no further
	while(l && zOutputFilename[l] != '.')	
	{
		l--;
	}

	//If the int l wasn't sucessful then get the zOutputFilename - 1
	if(!l)	
	{
		l = lstrlen(zOutputFilename) - 1;
	}


	zOutputFilename[l] = '\0';

	//Now add the file extension to the output string
	lstrcat(zOutputFilename, ".mp3");

	//Now Open The outputfile. meaning create it as binry
	// truncated, and let it accept the buffer to write to
	int hOut = open(zOutputFilename, O_WRONLY | O_BINARY | O_TRUNC | O_CREAT, S_IWRITE);

	//If Openoutput failed return a -1
	if(hOut == -1)	
	{
		return -1;
	}

	//getthe input wave file information
	GetInWaveProp(lpszWavFile);

	//Now import Lame's DLL Configuer class

	memset ( &b, 0, sizeof(b) );
	b.dwConfig = BE_CONFIG_LAME;
	b.format.LHV1.dwStructVersion = 1;
	b.format.LHV1.dwStructSize = sizeof(b);
	b.format.LHV1.dwSampleRate = myfreq;
	b.format.LHV1.dwReSampleRate = myfreq;
	b.format.LHV1.nMode = 0; 
	b.format.LHV1.dwBitrate = thisBitrate;
	b.format.LHV1.dwMpegVersion = MPEG1; // mpeg version (I or II)
	b.format.LHV1.dwPsyModel = 0;     // use default psychoacoustic model 
	b.format.LHV1.dwEmphasis = 0;     // no emphasis
	b.format.LHV1.bOriginal = TRUE;
	b.format.LHV1.dwMaxBitrate = thisBitrate;
	b.format.LHV1.bNoRes = TRUE;
	b.format.LHV1.nPreset = 0;
	b.format.LHV1.bCRC = thiscrc;
	b.format.LHV1.bCopyright = thiscopyright;
	b.format.LHV1.bOriginal = thisoriginal;
	b.format.LHV1.bPrivate = thisprivate;



    //set variables that will be used to hold a value for stream and buffer  etc etc :) 
	DWORD		dwSamples, dwMP3Buffer;
	HBE_STREAM	hbeStream;
	BE_ERR		err;

	//int err = start the stream by getting configuration, samples, buffer, and stream
	err = beInitStream( &b, &dwSamples, &dwMP3BufferSize, &hbeStream );

	//if it didnt get that info correctly return a -1
	if(err != BE_ERR_SUCCESSFUL)
	{
		return -1;
	}

	//set pMP3Buffer to new BYTE buffer
	PBYTE pMP3Buffer = new BYTE[dwMP3Buffer];

	//set pBuffer to new shrt samples
	PSHORT pBuffer = new SHORT[dwSamples];

	//if any of these conversions didnt work then return -1
	if(!pMP3Buffer || !pBuffer)
	{
		return -1;
	}

	//set variables for use in your code
	DWORD	length = filelength(hIn);
	DWORD	done = 0;
	DWORD	dwWrite;
	DWORD	toread;
	DWORD	towrite;
	
	//set buffer to the file stdout, character is set to NULL
	setbuf(stdout,NULL);

	//Now this loop says if there is still a buffer from the stream left then lets encode it
	while(done < length)
	{
		//set the variable still encoding to 1
		encoding_still = 1;

		//set up how much to readinto the buffer
		if(done + dwSamples * 2 < length)
		{

			toread = dwSamples * 2;
		}
		else
		{

			toread = length - done;
		}


		//if there is nothing left in the buffer to read return a -1
		if(read(hIn, pBuffer, toread) == -1)
		{
			encoding_still = 0;
			return -1;
		}		 

		//now after you've read a chunk of the wave file into the buffer start encoding the stream in teh buffer
		err = beEncodeChunk(hbeStream, toread/2, pBuffer, pMP3Buffer, &towrite);

		//if there is nothing in the buffer left to encode then return -1
		if(err != BE_ERR_SUCCESSFUL)
		{
			//close everything out, set still encoding to 0 and return a -1
			beCloseStream(hbeStream);
			encoding_still = 0;
			return -1;
		}
		
		//if there is no data left to write close it out
		if(write(hOut, pMP3Buffer, towrite) == -1)
		{
			encoding_still = 0;
			return -1;
		}

		//if cancel was pressed encode the rest of the buffer then exit out
		if (cancel == TRUE)
		{
			beCloseStream(hbeStream);
			close(hIn);
			close(hOut);

			if (hLameDLL)
				FreeLibrary(hLameDLL);

			return 0;
		}

		//how much was read
		done += toread;
		
		//figure percentage how much was done / length * 100
		percent_done = 100 * (float)done/(float)length;

		// call the enumerated function to display
		// completion of encoding process
		if ((EnumEncoding((int)percent_done)) == FALSE)
			return -2;	//indicate that encoding was stopped by user
	}
	//Close out everything
	encoding_still = 0;

	err = beDeinitStream(hbeStream, pMP3Buffer, &dwWrite);
	//if close out was unsuccessful mnually close stream and return -1
	if(err != BE_ERR_SUCCESSFUL)
	{
		beCloseStream(hbeStream);
		return -1;
	}

	//if nothing left to write close out writeing to the output file
	if(dwWrite)
	{
		if(write(hOut, pMP3Buffer, dwWrite) == -1)
		{
			return -1;
		}
	}

	//close stream
	beCloseStream(hbeStream);

	//clseoutput and input file
	close(hIn);
	close(hOut);

	//recet percent done for prog bar
	percent_done = 0;

	// cleanup processing
	if (hLameDLL)
		FreeLibrary(hLameDLL);


	return TRUE;
}