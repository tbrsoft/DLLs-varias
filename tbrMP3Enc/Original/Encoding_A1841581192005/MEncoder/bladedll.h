

#ifndef ___BLADEDLL_H_INCLUDED___

#define ___BLADEDLL_H_INCLUDED___

#pragma pack(push)
#pragma pack(1)

#ifndef __GNUC__
#define PACKED
#endif

/* encoding formats */
#define         BE_CONFIG_MP3          0
#define         BE_CONFIG_LAME         256
#define         BE_CONFIG_VORBIS       512


/* type definitions */
typedef         unsigned long          HBE_STREAM;
typedef         HBE_STREAM            *PHBE_STREAM;
typedef         unsigned long          BE_ERR;


/* error codes */
#define         BE_ERR_SUCCESSFUL      0x00000000
#define         BE_ERR_INVALID_FORMAT  0x00000001
#define         BE_ERR_INVALID_FORMAT_PARAMETERS 0x00000002
#define         BE_ERR_NO_MORE_HANDLES 0x00000003
#define         BE_ERR_INVALID_HANDLE  0x00000004
#define         BE_ERR_BUFFER_TOO_SMALL 0x00000005


/* other constants */
#define		BE_MAX_HOMEPAGE			256


/* format specific variables */
#define  BE_MP3_MODE_STEREO          0
#define  BE_MP3_MODE_JSTEREO         1
#define  BE_MP3_MODE_DUALCHANNEL     2
#define  BE_MP3_MODE_MONO            3


#define  MPEG1    1
#define  MPEG2    0


/* vorb_enc.dll struct version */
#define  VORBENCSTRUCTVER             2


typedef enum 
{
  NORMAL_QUALITY=0,
  LOW_QUALITY,
  HIGH_QUALITY,
  VOICE_QUALITY,
} MPEG_QUALITY;


typedef enum
{
	VBR_METHOD_NONE			= -1,
	VBR_METHOD_DEFAULT		=  0,
	VBR_METHOD_OLD			=  1,
	VBR_METHOD_NEW			=  2,
	VBR_METHOD_MTRH			=  3,
	VBR_METHOD_ABR			=  4
} VBRMETHOD;


typedef enum 
{
	LQP_NOPRESET			=-1,

	// QUALITY PRESETS
	LQP_NORMAL_QUALITY		= 0,
	LQP_LOW_QUALITY			= 1,
	LQP_HIGH_QUALITY		= 2,
	LQP_VOICE_QUALITY		= 3,
	LQP_R3MIX				= 4,
	LQP_VERYHIGH_QUALITY	= 5,
	LQP_STANDARD			= 6,
	LQP_FAST_STANDARD		= 7,
	LQP_EXTREME				= 8,
	LQP_FAST_EXTREME		= 9,
	LQP_INSANE				= 10,
	LQP_ABR					= 11,
	LQP_CBR					= 12,
	LQP_MEDIUM				= 13,
	LQP_FAST_MEDIUM			= 14,

	// NEW PRESET VALUES
	LQP_PHONE	=1000,
	LQP_SW		=2000,
	LQP_AM		=3000,
	LQP_FM		=4000,
	LQP_VOICE	=5000,
	LQP_RADIO	=6000,
	LQP_TAPE	=7000,
	LQP_HIFI	=8000,
	LQP_CD		=9000,
	LQP_STUDIO	=10000

} LAME_QUALITY_PRESET;


typedef struct	
{
  DWORD	dwConfig;              // BE_CONFIG_XXXXX
  union {
    struct {  
      DWORD  dwSampleRate;     // 48000, 44100 and 32000 allowed
      BYTE   byMode;           // BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL,
                               // BE_MP3_MODE_MONO
      WORD   wBitrate;         // 32, 40, 48, 56, 64, 80, 96, 112, 128,
                               // 160, 192, 224, 256 and 320 allowed
      BOOL   bPrivate;		
      BOOL   bCRC;
      BOOL   bCopyright;
      BOOL   bOriginal;
    } PACKED mp3;  // BE_CONFIG_MP3

    struct {

      // STRUCTURE INFORMATION
      DWORD  dwStructVersion;
      DWORD  dwStructSize;

      // BASIC ENCODER SETTINGS
	  DWORD			dwSampleRate;		// SAMPLERATE OF INPUT FILE
	  DWORD			dwReSampleRate;		// DOWNSAMPLERATE, 0=ENCODER DECIDES  
	  LONG			nMode;				// BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL, BE_MP3_MODE_MONO
	  DWORD			dwBitrate;			// CBR bitrate, VBR min bitrate
	  DWORD			dwMaxBitrate;		// CBR ignored, VBR Max bitrate
	  LONG			nPreset;			// Quality preset, use one of the settings of the LAME_QUALITY_PRESET enum
	  DWORD			dwMpegVersion;		// FUTURE USE, MPEG-1 OR MPEG-2
	  DWORD			dwPsyModel;			// FUTURE USE, SET TO 0
	  DWORD			dwEmphasis;			// FUTURE USE, SET TO 0


      // BIT STREAM SETTINGS
      BOOL   bPrivate;          // Set Private Bit
      BOOL   bCRC;              // Insert CRC
      BOOL   bCopyright;        // Set copyright bit
      BOOL   bOriginal;         // Set original bit

      // VBR STUFF
	  BOOL			bWriteVBRHeader;	// WRITE XING VBR HEADER (TRUE/FALSE)
	  BOOL			bEnableVBR;			// USE VBR ENCODING (TRUE/FALSE)
	  INT			nVBRQuality;		// VBR QUALITY 0..9
	  DWORD			dwVbrAbr_bps;		// Use ABR in stead of nVBRQuality
	  VBRMETHOD		nVbrMethod;
	  BOOL			bNoRes;				// Disable Bit resorvoir (TRUE/FALSE)


	  // MISC SETTINGS
	  BOOL			bStrictIso;			// Use strict ISO encoding rules (TRUE/FALSE)
	  WORD			nQuality;			// Quality Setting, HIGH BYTE should be NOT LOW byte, otherwhise quality=5


      BYTE   btReserved[255];   // Reserved for future use
    } LHV1;

    struct {
      int version;              // set to VORBENCSTRUCTVER
      int channels;             // CD audio == 2
      long rate;                // CD audio == 44100

      char *szTitle;            // track title
      char *szVersion;          // used to designate mult. versions of same
                                // track
      char *szAlbum;            // Album name
      char *szArtist;           // Artist's name
      char *szOrganization;     // Organization (or record label)
      char *szDescription;      // short description of the track contents
      char *szGenre;            // text indication of the genre
      char *szDate;             // date the track was recorded
      char *szLocation;         // place where the track was recorded
      char *szCopyright;        // copyright info
      int mode;                 // sets which info_* to use 0 = A, ...
                                // According to Monty, the info_* structs
                                // are a hack, and will likely go away.
      long minbitrate;
      long maxbitrate;
      long nominalbitrate;
    } vorb;

    struct {
      DWORD  dwSampleRate;
      BYTE   byMode;
      WORD   wBitrate;
      BYTE   byEncodingMethod;
    } PACKED aac;
    
  } PACKED format;
  
} PACKED BE_CONFIG, *PBE_CONFIG;


typedef struct	
{
  // BladeEnc DLL Version number
  BYTE	byDLLMajorVersion;
  BYTE	byDLLMinorVersion;
  
  // BladeEnc Engine Version Number
  BYTE	byMajorVersion;
  BYTE	byMinorVersion;
  
  // DLL Release date
  BYTE	byDay;
  BYTE	byMonth;
  WORD	wYear;
  
  // BladeEnc	Homepage URL
  
  CHAR	zHomepage[BE_MAX_HOMEPAGE + 1];	
  
} PACKED BE_VERSION, *PBE_VERSION;			


#ifndef _BLADEDLL


typedef BE_ERR	(*BEINITSTREAM)		(PBE_CONFIG, PDWORD, PDWORD, PHBE_STREAM);
typedef BE_ERR	(*BEENCODECHUNK)	(HBE_STREAM, DWORD, PSHORT, PBYTE, PDWORD);
typedef BE_ERR	(*BEDEINITSTREAM)	(HBE_STREAM, PBYTE, PDWORD);
typedef BE_ERR	(*BECLOSESTREAM)	(HBE_STREAM);
typedef VOID	(*BEVERSION)		(PBE_VERSION);
typedef BE_ERR  (*BEWRITEVBRHEADER)     (LPCSTR);


#define	TEXT_BEINITSTREAM	"beInitStream"
#define	TEXT_BEENCODECHUNK	"beEncodeChunk"
#define	TEXT_BEDEINITSTREAM	"beDeinitStream"
#define	TEXT_BECLOSESTREAM	"beCloseStream"
#define	TEXT_BEVERSION		"beVersion"
#define TEXT_BEWRITEVBRHEADER   "beWriteVBRHeader"

	
#else


__declspec(dllexport) BE_ERR	beInitStream(PBE_CONFIG pbeConfig, PDWORD dwSamples, PDWORD dwBufferSize, PHBE_STREAM phbeStream);
__declspec(dllexport) BE_ERR	beEncodeChunk(HBE_STREAM hbeStream, DWORD nSamples, PSHORT pSamples, PBYTE pOutput, PDWORD pdwOutput);
__declspec(dllexport) BE_ERR	beDeinitStream(HBE_STREAM hbeStream, PBYTE pOutput, PDWORD pdwOutput);
__declspec(dllexport) BE_ERR	beCloseStream(HBE_STREAM hbeStream);
__declspec(dllexport) VOID		beVersion(PBE_VERSION pbeVersion);


#endif


#pragma pack(pop)


#endif
