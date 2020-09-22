Attribute VB_Name = "Id3Module"
Public Type Id3                 'This type is standard for
    Title As String * 30            ' Id3 Tags
    Artist As String * 30           ' Although later versions
    Album As String * 30            ' use comments for 28 bytes
    sYear  As String * 4            ' and they use the 2 remaining  bytes for "TrackNumber"!
    Comments As String * 30
    Genre As Byte
End Type
Public id3Info As Id3           ' Declare a variable as the id3 type
Public GenreArray() As String         ' we use this array to fill all the Genre's ( look in form load)
Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
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
Public Function GetId3(FileName As String) As Boolean
Dim TaG As String * 3  ' We use this variable to make sure the file has an ID3TAG
Open FileName For Binary As #1
    Get #1, FileLen(FileName) - 127, TaG
    If TaG = "TAG" Then
        Get #1, FileLen(FileName) - 124, id3Info
        GetId3 = True
    Else
        GetId3 = False
    End If
Close #1
End Function
Public Function SaveId3(FileName As String, MP3Info As Id3)
Dim TaG As String * 3
Open FileName For Binary As #1
    Get #1, FileLen(FileName) - 127, TaG
    If TaG = "TAG" Then
        Put #1, FileLen(FileName) - 124, MP3Info
    Else
        Put #1, FileLen(FileName) + 1, "TAG"
        Put #1, FileLen(FileName) + 4, MP3Info
        Close #1
    End If
Close #1
End Function
