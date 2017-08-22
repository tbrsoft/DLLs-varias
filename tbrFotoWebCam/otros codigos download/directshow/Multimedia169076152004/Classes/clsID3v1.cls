VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ID3v1 Class Module by Vesa Piittinen aka Merri
' http://merri.net
'
' remember to close file before saving ID3v1 information


Option Explicit


'helps loading and saving data
Private Type ID3v1_TAG              '(128 bytes)
    Tag As String * 3               'always TAG
    Title As String * 30            'title, 30 characters
    Artist As String * 30           'artist, 30 characters
    Album As String * 30            'album, 30 characters
    Year As String * 4              'year, 4 characters
    Comment As String * 30          'comment, 30 characters (or 28 if track# included)
    Genre As Byte                   'genre, 255 for none defined
End Type

'storaging data for use
Private Type ID3v1_STORE
    Title As String                 'title
    Artist As String                'artist
    Album As String                 'album
    Year As Integer                 'year
    Comment As String               'comment
    Genre As Byte                   'genre (use GetGenre for text)
    Track As Byte                   'track
End Type


Private ID3v1genres(255) As String              'ID3v1 genrelist storage for fast access
Private ID3v1store As ID3v1_STORE               'ID3v1 tag information


'ID3v1 class initialize
Private Sub Class_Initialize()
    'initialize genres for fast access
    'based on WinAmp genres
    ID3v1genres(147) = "Synthpop"
    ID3v1genres(146) = "JPop"
    ID3v1genres(145) = "Anime"
    ID3v1genres(144) = "Thrash Metal"
    ID3v1genres(143) = "Salsa"
    ID3v1genres(142) = "Merengue"
    ID3v1genres(141) = "Christian Rock"
    ID3v1genres(140) = "Contemporary Christian"
    ID3v1genres(139) = "Crossover"
    ID3v1genres(138) = "Black Metal"
    ID3v1genres(137) = "Heavy Metal"
    ID3v1genres(136) = "Christian Gangsta Rap"
    ID3v1genres(135) = "Beat"
    ID3v1genres(134) = "Polsk Punk"
    ID3v1genres(133) = "Negerpunk"
    ID3v1genres(132) = "BritPop"
    ID3v1genres(131) = "Indie"
    ID3v1genres(130) = "Terror"
    ID3v1genres(129) = "Hardcore"
    ID3v1genres(128) = "Club-House"
    ID3v1genres(127) = "Drum & Bass"
    ID3v1genres(126) = "Goa"
    ID3v1genres(125) = "Dance Hall"
    ID3v1genres(124) = "Euro-House"
    ID3v1genres(123) = "A Cappella"
    ID3v1genres(122) = "Drum Solo"
    ID3v1genres(121) = "Punk Rock"
    ID3v1genres(120) = "Duet"
    ID3v1genres(119) = "Freestyle"
    ID3v1genres(118) = "Rhythmic Soul"
    ID3v1genres(117) = "Power Ballad"
    ID3v1genres(116) = "Ballad"
    ID3v1genres(115) = "Folklore"
    ID3v1genres(114) = "Samba"
    ID3v1genres(113) = "Tango"
    ID3v1genres(112) = "Club"
    ID3v1genres(111) = "Slow Jam"
    ID3v1genres(110) = "Satire"
    ID3v1genres(109) = "Porn Groove"
    ID3v1genres(108) = "Primus"
    ID3v1genres(107) = "Booty Bass"
    ID3v1genres(106) = "Symphony"
    ID3v1genres(105) = "Sonata"
    ID3v1genres(104) = "Chamber Music"
    ID3v1genres(103) = "Opera"
    ID3v1genres(102) = "Chanson"
    ID3v1genres(101) = "Speech"
    ID3v1genres(100) = "Humour"
    ID3v1genres(99) = "Acoustic"
    ID3v1genres(98) = "Easy Listening"
    ID3v1genres(97) = "Chorus"
    ID3v1genres(96) = "Big Band"
    ID3v1genres(95) = "Slow Rock"
    ID3v1genres(94) = "Symphonic Rock"
    ID3v1genres(93) = "Psychedelic Rock"
    ID3v1genres(92) = "Progressive Rock"
    ID3v1genres(91) = "Gothic Rock"
    ID3v1genres(90) = "Avantgarde"
    ID3v1genres(89) = "Bluegrass"
    ID3v1genres(88) = "Celtic"
    ID3v1genres(87) = "Revival"
    ID3v1genres(86) = "Latin"
    ID3v1genres(85) = "Bebob"
    ID3v1genres(84) = "Fast-Fusion"
    ID3v1genres(83) = "Swing"
    ID3v1genres(82) = "National Folk"
    ID3v1genres(81) = "Folk/Rock"
    ID3v1genres(80) = "Folk"
    ID3v1genres(79) = "Hard Rock"
    ID3v1genres(78) = "Rock & Roll"
    ID3v1genres(77) = "Musical"
    ID3v1genres(76) = "Retro"
    ID3v1genres(75) = "Polka"
    ID3v1genres(74) = "Acid Jazz"
    ID3v1genres(73) = "Acid Punk"
    ID3v1genres(72) = "Tribal"
    ID3v1genres(71) = "Lo-Fi"
    ID3v1genres(70) = "Trailer"
    ID3v1genres(69) = "Showtunes"
    ID3v1genres(68) = "Rave"
    ID3v1genres(67) = "Psychedelic"
    ID3v1genres(66) = "New Wave"
    ID3v1genres(65) = "Cabaret"
    ID3v1genres(64) = "Native American"
    ID3v1genres(63) = "Jungle"
    ID3v1genres(62) = "Pop/Funk"
    ID3v1genres(61) = "Christian Rap"
    ID3v1genres(60) = "Top 40"
    ID3v1genres(59) = "Gangsta Rap"
    ID3v1genres(58) = "Cult"
    ID3v1genres(57) = "Comedy"
    ID3v1genres(56) = "Southern Rock"
    ID3v1genres(55) = "Dream"
    ID3v1genres(54) = "Eurodance"
    ID3v1genres(53) = "Pop-Folk"
    ID3v1genres(52) = "Electronic"
    ID3v1genres(51) = "Techno-Industrial"
    ID3v1genres(50) = "Darkwave"
    ID3v1genres(49) = "Gothic"
    ID3v1genres(48) = "Ethnic"
    ID3v1genres(47) = "Instrumental Rock"
    ID3v1genres(46) = "Instrumental Pop"
    ID3v1genres(45) = "Meditative"
    ID3v1genres(44) = "Space"
    ID3v1genres(43) = "Punk"
    ID3v1genres(42) = "Soul"
    ID3v1genres(41) = "Bass"
    ID3v1genres(40) = "Alt. Rock"
    ID3v1genres(39) = "Noise"
    ID3v1genres(38) = "Gospel"
    ID3v1genres(37) = "Sound Clip"
    ID3v1genres(36) = "Game"
    ID3v1genres(35) = "House"
    ID3v1genres(34) = "Acid"
    ID3v1genres(33) = "Instrumental"
    ID3v1genres(32) = "Classical"
    ID3v1genres(31) = "Trance"
    ID3v1genres(30) = "Fusion"
    ID3v1genres(29) = "Jazz+Funk"
    ID3v1genres(28) = "Vocal"
    ID3v1genres(27) = "Trip-Hop"
    ID3v1genres(26) = "Ambient"
    ID3v1genres(25) = "Euro-Techno"
    ID3v1genres(24) = "Soundtrack"
    ID3v1genres(23) = "Pranks"
    ID3v1genres(22) = "Death Metal"
    ID3v1genres(21) = "Ska"
    ID3v1genres(20) = "Alternative"
    ID3v1genres(19) = "Industrial"
    ID3v1genres(18) = "Techno"
    ID3v1genres(17) = "Rock"
    ID3v1genres(16) = "Reggae"
    ID3v1genres(15) = "Rap"
    ID3v1genres(14) = "R&B"
    ID3v1genres(13) = "Pop"
    ID3v1genres(12) = "Other"
    ID3v1genres(11) = "Oldies"
    ID3v1genres(10) = "New Age"
    ID3v1genres(9) = "Metal"
    ID3v1genres(8) = "Jazz"
    ID3v1genres(7) = "Hip-Hop"
    ID3v1genres(6) = "Grunge"
    ID3v1genres(5) = "Funk"
    ID3v1genres(4) = "Disco"
    ID3v1genres(3) = "Dance"
    ID3v1genres(2) = "Country"
    ID3v1genres(1) = "Classic Rock"
    ID3v1genres(0) = "Blues"
End Sub


'ID3v1 class information functions
Public Function Album() As String
    Album = ID3v1store.Album
End Function
Public Function Artist() As String
    Artist = ID3v1store.Artist
End Function
Public Function Comment() As String
    Comment = ID3v1store.Comment
End Function
Public Function Genre() As Byte
    Genre = ID3v1store.Genre
End Function
Public Sub GetTag(ByRef tTitle As String, ByRef tArtist As String, Optional ByRef tAlbum As String, Optional ByRef tYear As Integer, Optional ByRef tComment As String, Optional ByRef tTrack As Byte, Optional ByRef tGenre As Byte)
    'return tag information
    With ID3v1store
        tTitle = .Title
        tArtist = .Artist
        If Not IsMissing(tAlbum) Then tAlbum = .Album
        If Not IsMissing(tYear) Then tYear = .Year
        If Not IsMissing(tComment) Then tComment = .Comment
        If Not IsMissing(tTrack) Then tTrack = .Track
        If Not IsMissing(tGenre) Then tGenre = .Genre
    End With
End Sub
Public Function StrGenre() As String
    StrGenre = ID3v1genres(ID3v1store.Genre)
End Function
Public Function Title() As String
    Title = ID3v1store.Title
End Function
Public Function Track() As Byte
    Track = ID3v1store.Track
End Function
Public Function Year() As Integer
    Year = ID3v1store.Year
End Function


'ID3v1 class file functions
Public Function OpenFile(ByVal File As String) As Boolean
    Dim IDtag As ID3v1_TAG, FileNumber As Byte
    On Error GoTo ErrorHandler
    FileNumber = FreeFile
    'open file for read
    Open File For Binary Access Read As #FileNumber
        'check if file is big enough for it to contain a tag
        If LOF(1) < Len(IDtag) Then Close #FileNumber: Exit Function
        'read the tag
        Get #FileNumber, LOF(1) - Len(IDtag) + 1, IDtag
    Close #FileNumber
    'check if there is a tag
    If IDtag.Tag <> "TAG" Then
        'no tag, clear all info
        With ID3v1store
            .Album = ""
            .Artist = ""
            .Title = ""
            .Year = 0
            .Comment = ""
            .Track = 0
            .Genre = 255 '255 is reserved for none
        End With
        Exit Function
    End If
    'start reading and formatting the information
    With ID3v1store
        .Album = NullTrim(IDtag.Album)
        .Artist = NullTrim(IDtag.Artist)
        .Title = NullTrim(IDtag.Title)
        If IsNumeric(IDtag.Year) Then
            .Year = CInt(IDtag.Year)
        Else
            .Year = 0
        End If
        .Genre = IDtag.Genre
        'check if there is track information (byte before the last byte is zero if track information exists)
        If Asc(Mid$(IDtag.Comment, 29, 1)) = 0 Then
            .Comment = NullTrim(Left$(IDtag.Comment, 28))
            'convert character to character code
            .Track = Asc(Right$(IDtag.Comment, 1))
        Else
            .Comment = NullTrim(IDtag.Comment)
            .Track = 0
        End If
    End With
    'success!
    OpenFile = True
    Exit Function
ErrorHandler:
    Close #FileNumber
End Function
Public Function SaveFile(ByVal File As String, ByVal Title As String, ByVal Artist As String, ByVal Album As String, ByVal Year As Integer, ByVal Comment As String, ByVal Track As Byte, ByVal Genre As Byte) As Boolean
    Dim WriteTag As ID3v1_TAG, ReadTag As ID3v1_TAG, FileNumber As Byte
    On Error GoTo ErrorHandler
    FileNumber = FreeFile
    'error correction
    If Len(Title) > 30 Then Title = Left$(Title, 30)
    If Len(Artist) > 30 Then Artist = Left$(Artist, 30)
    If Len(Album) > 30 Then Album = Left$(Album, 30)
    If Year > Val(Format(Date, "yyyy")) Then Year = Val(Format(Date, "yyyy"))
    If Len(Comment) > 30 Then Comment = Left$(Comment, 30)
    'set up tag to be written
    WriteTag.Tag = "TAG"
    WriteTag.Title = Title
    WriteTag.Artist = Artist
    WriteTag.Album = Album
    WriteTag.Year = Format$(Year, "0000")
    WriteTag.Comment = Comment
    'if there is track information to be written...
    If Track <> 0 Then WriteTag.Comment = Left$(WriteTag.Comment, 28) & vbNullChar & Chr$(Track)
    WriteTag.Genre = Genre
    'open file for processing
    Open File For Binary As #FileNumber
        If LOF(1) < Len(ReadTag) Then Close #FileNumber: Exit Function
        'read a tag
        Get #FileNumber, LOF(1) - Len(ReadTag) + 1, ReadTag
        'check if there is a tag
        If ReadTag.Tag = "TAG" Then
            'overwrite old tag
            Put #FileNumber, LOF(1) - Len(WriteTag) + 1, WriteTag
        Else
            'go to end of file
            Seek #FileNumber, LOF(1) + 1
            'add tag
            Put #FileNumber, , WriteTag
        End If
    Close #FileNumber
    'success!
    SaveFile = True
    Exit Function
ErrorHandler:
    Close #FileNumber
End Function


'ID3v1 class additional functions
Private Function NullTrim(ByVal Text As String) As String
    'trim extra nulls away (most often in the end of the string) keeping unicode support
    'Do While InStr(Temp, vbNullChar & vbNullChar)
    '    Temp = Replace$(Temp, vbNullChar & vbNullChar, vbNullChar)
    'Loop
    'trim
    NullTrim = Trim$(Replace$(Text, vbNullChar, vbNullString))
End Function
