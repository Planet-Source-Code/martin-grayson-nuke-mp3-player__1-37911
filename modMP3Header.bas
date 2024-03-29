Attribute VB_Name = "modMP3Header"


Option Explicit
Public sGenreMatrix

Type Info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sComment As String * 30
    sYear As String * 4
    sGenre As String * 21
End Type

Type HeaderInfo
    Layer As String
    Frequency As String
    Bitrate As String
    Mode As String
    MpegVersion As String
    Emphasis As String
    FPlayTime As String
    mFileSize As String
End Type

Public MP3Info As Info
Public MP3HeaderInfo As HeaderInfo

Public Function GetMP3Tag(ByVal sPassFileName As String) As Boolean
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    Dim sGenreMatrix As String
    Dim sGenre() As String
    
 
    sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

    sGenre = Split(sGenreMatrix, "|")
    

    MP3Info.sTitle = ""
    MP3Info.sArtist = ""
    MP3Info.sAlbum = ""
    MP3Info.sYear = ""
    MP3Info.sComment = ""

    If Dir(sPassFileName) = "" Then GetMP3Tag = False: GoTo CloseMe
 
    GetMP3Tag = True
    iFreefile = FreeFile
    lFilePos = FileLen(sPassFileName) - 127
    Open sPassFileName For Binary As #iFreefile
    Get #iFreefile, lFilePos, sData
    Close #iFreefile

    
    If Left(sData, 3) = "TAG" Then
        MP3Info.sTitle = RTrim(Mid(sData, 4, 30))
        MP3Info.sArtist = RTrim(Mid(sData, 34, 30))
        MP3Info.sAlbum = RTrim(Mid(sData, 64, 30))
        MP3Info.sYear = RTrim(Mid(sData, 94, 4))
        MP3Info.sComment = RTrim(Mid(sData, 98, 30))
        MP3Info.sGenre = RTrim(sGenre(Asc(Mid(sData, 128, 1))))
    End If
    
CloseMe:
Close #iFreefile
End Function



Public Function ReadMP3Header(sPassFileName As String)
Dim z, i
Dim BinaryString As String
Dim byteArray(4) As Byte
Dim bin As String
Dim BinString As String
Dim DecString As Integer

Open sPassFileName For Binary Access Read As #1
   For z = 1 To 4
   Get #1, z, byteArray(z)
   Next z
 Close #1
 bin = ""
   For z = 1 To 4
     For i = 0 To 7 Step 1
         If byteArray(z) And (2 ^ i) Then
            bin = bin + "1"
            Else
            bin = bin + "0"
         End If
         Next i
Next z
BinaryString = bin

DecString = 0
BinString = Mid(bin, 19, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Frequency = 44100
Case 1
MP3HeaderInfo.Frequency = 32000
Case 2
MP3HeaderInfo.Frequency = 48000
Case 3
End Select

DecString = 0
BinString = Mid(bin, 10, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Layer = ""
Case 1
MP3HeaderInfo.Layer = 2
Case 2
MP3HeaderInfo.Layer = 3
Case 3
MP3HeaderInfo.Layer = 1
End Select

DecString = 0
BinString = Mid(bin, 31, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Mode = "Stereo"
Case 1
MP3HeaderInfo.Mode = "Dual Channel"
Case 2
MP3HeaderInfo.Mode = "Joint stereo"
Case 3
MP3HeaderInfo.Mode = "Mono"
End Select

If Mid(bin, 12, 1) = 0 Then
MP3HeaderInfo.MpegVersion = 2
Else
MP3HeaderInfo.MpegVersion = 1
End If

DecString = 0
BinString = Mid(bin, 21, 4)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Bitrate = 0
Case 1
MP3HeaderInfo.Bitrate = 112
Case 2
MP3HeaderInfo.Bitrate = 56
Case 3
MP3HeaderInfo.Bitrate = 224
Case 4
MP3HeaderInfo.Bitrate = 40
Case 5
MP3HeaderInfo.Bitrate = 160
Case 6
MP3HeaderInfo.Bitrate = 80
Case 7
MP3HeaderInfo.Bitrate = 320
Case 8
MP3HeaderInfo.Bitrate = 32
Case 9
MP3HeaderInfo.Bitrate = 128
Case 10
MP3HeaderInfo.Bitrate = 64
Case 11
MP3HeaderInfo.Bitrate = 256
Case 12
MP3HeaderInfo.Bitrate = 48
Case 13
MP3HeaderInfo.Bitrate = 192
Case 14
MP3HeaderInfo.Bitrate = 96
Case 15
MP3HeaderInfo.Bitrate = 0
If MP3HeaderInfo.Layer = 1 Then
    Select Case DecString
    Case 0
MP3HeaderInfo.Bitrate = 0
    Case 1
  MP3HeaderInfo.Bitrate = 128
    Case 2
   MP3HeaderInfo.Bitrate = 64
    Case 3
MP3HeaderInfo.Bitrate = 256
    Case 4
MP3HeaderInfo.Bitrate = 48
    Case 5
MP3HeaderInfo.Bitrate = 192
    Case 6
MP3HeaderInfo.Bitrate = 96
    Case 7
    MP3HeaderInfo.Bitrate = 384
    Case 8
MP3HeaderInfo.Bitrate = 32
    Case 9
MP3HeaderInfo.Bitrate = 160
    Case 10
    MP3HeaderInfo.Bitrate = 80
    Case 11
MP3HeaderInfo.Bitrate = 320
    Case 12
MP3HeaderInfo.Bitrate = 56
    Case 13
MP3HeaderInfo.Bitrate = 224
    Case 14
  MP3HeaderInfo.Bitrate = 112
    Case 15
MP3HeaderInfo.Bitrate = 0
End Select
End If
End Select

DecString = 0
BinString = Mid(bin, 25, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Emphasis = "No"
Case 1
MP3HeaderInfo.Emphasis = "-?-"
Case 2
MP3HeaderInfo.Emphasis = "50/15"
Case 3
MP3HeaderInfo.Emphasis = "CITT j. 17"
End Select

With MP3HeaderInfo
    Dim min, sec
    .Bitrate = Int(.Bitrate)
    .mFileSize = FileSizeMP3(sPassFileName)
    .FPlayTime = ((.mFileSize * 8) / (.Bitrate * 1000))
    min = .FPlayTime \ 60
    sec = .FPlayTime - (min * 60)
    .FPlayTime = Format(min, "#0#") & ":" & Format(sec, "0#")
End With

End Function

Public Function RemoveMP3Tag(sPassFileName As String) As Boolean
Dim blank
On Error GoTo Errorcheck
blank = String$(127, 0)
Open sPassFileName For Binary Access Write As #1
Seek #1, LOF(1) - 127
Put #1, , blank
Close #1
RemoveMP3Tag = True
Exit Function
Errorcheck:
RemoveMP3Tag = False
Exit Function
End Function

Public Function WriteMP3Tag(sPassFileName As String, TAG As String, SongName As String, Artist As String, Album As String, Year As String, Comment As String, Genre As String) As Boolean
On Error GoTo Errorcheck
 Dim wTag As String * 3
 Dim wSongname As String * 30
 Dim wArtist As String * 30
 Dim wAlbum As String * 30
 Dim wYear As String * 4
 Dim wComment As String * 30
 Dim wGenre As String * 1
 
    wTag = TAG
    wSongname = RTrim(SongName)
    wArtist = RTrim(Artist)
    wAlbum = RTrim(Album)
    wYear = Left(Year, 4)
    wComment = RTrim(Comment)
    wGenre = Chr(Genre + 1)
    
    Open sPassFileName For Binary Access Write As #1
    Seek #1, FileLen(sPassFileName) - 127
    Put #1, , wTag
    Put #1, , wSongname
    Put #1, , wArtist
    Put #1, , wAlbum
    Put #1, , wYear
    Put #1, , wComment
    Put #1, , wGenre
    WriteMP3Tag = True
        Close #1
Exit Function
Errorcheck:
    WriteMP3Tag = False
Exit Function
End Function

Public Function FileSizeMP3(File As String) As String
    Dim LSize As String
    If File = "" Then
    FileSizeMP3 = ""
    Exit Function
    End If
    LSize = FileLen(File)
    FileSizeMP3 = LSize
End Function
