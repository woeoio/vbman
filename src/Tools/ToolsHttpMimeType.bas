Attribute VB_Name = "ToolsHttpMimeType"
Option Explicit

' Modify by woeoio@qq.com
' Form Jason Peter Brown's VbWebserver

' Copyright (c) 2017 Jason Peter Brown <jason@bitspaces.com>
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


Public Function MapMimeType(FilePath As String) As String
    Dim l_TypeInfo As String, LastPart As String, Arr, Max As Long
    Arr = Split(FilePath, ".")
    Max = UBound(Arr)
    LastPart = Arr(Max)
    Select Case LastPart
    Case "7z"
        ' 7-Zip
        l_TypeInfo = "application/x-7z-compressed"
    Case "aac"
        ' Advanced Audio Coding (AAC)
        l_TypeInfo = "audio/x-aac"
    Case "avi"
        ' Audio Video Interleave (AVI)
        l_TypeInfo = "video/x-msvideo"
    Case "bmp"
        ' Bitmap Image File
        l_TypeInfo = "image/bmp"
    Case "css"
        ' Cascading Style Sheets (CSS)
        l_TypeInfo = "text/css"
    Case "csv"
        ' Comma-Seperated Values
        l_TypeInfo = "text/csv"
    Case "doc"
        ' Microsoft Word
        l_TypeInfo = "application/msword"
    Case "docx"
        ' Microsoft Office - OOXML - Word Document
        l_TypeInfo = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    Case "dwf"
        ' Autodesk Design Web Format (DWF)
        l_TypeInfo = "model/vnd.dwf"
    Case "dwg"
        ' DWG Drawing
        l_TypeInfo = "image/vnd.dwg"
    Case "dxf"
        ' AutoCAD DXF
        l_TypeInfo = "image/vnd.dxf"
    Case "eml"
        ' Email Message
        l_TypeInfo = "message/rfc822"
    Case "exe"
        ' Microsoft Application
        l_TypeInfo = "application/x-msdownload"
    Case "f4v"
        ' Flash Video
        l_TypeInfo = "video/x-f4v"
    Case "flv"
        ' Flash Video
        l_TypeInfo = "video/x-flv"
    Case "gif"
        ' Graphics Interchange Format
        l_TypeInfo = "image/gif"
    Case "html"
        ' HyperText Markup Language (HTML)
        l_TypeInfo = "text/html"
    Case "ico"
        ' Icon Image
        l_TypeInfo = "image/x-icon"
    Case "ics"
        ' iCalendar
        l_TypeInfo = "text/calendar"
    Case "jar"
        ' Java Archive
        l_TypeInfo = "application/java-archive"
    Case "java"
        ' Java Source File
        l_TypeInfo = "text/x-java-source,java"
    Case "jpeg"
        ' JPEG Image
        l_TypeInfo = "image/jpeg"
    Case "jpg"
        ' JPEG Image
        l_TypeInfo = "image/jpeg"
    Case "jpgv"
        ' JPGVideo
        l_TypeInfo = "video/jpeg"
    Case "jpm"
        ' JPEG 2000 Compound Image File Format
        l_TypeInfo = "video/jpm"
    Case "js"
        ' JavaScript
        l_TypeInfo = "application/javascript"
    Case "json"
        ' JavaScript Object Notation (JSON)
        l_TypeInfo = "application/json"
    Case "kml"
        ' Google Earth - KML
        l_TypeInfo = "application/vnd.google-earth.kml+xml"
    Case "kmz"
        ' Google Earth - Zipped KML
        l_TypeInfo = "application/vnd.google-earth.kmz"
    Case "m3u"
        ' M3U (Multimedia Playlist)
        l_TypeInfo = "audio/x-mpegurl"
    Case "m3u8"
        ' Multimedia Playlist Unicode
        l_TypeInfo = "application/vnd.apple.mpegurl"
    Case "mdb"
        ' Microsoft Access
        l_TypeInfo = "application/x-msaccess"
    Case "mid"
        ' MIDI - Musical Instrument Digital Interface
        l_TypeInfo = "audio/midi"
    Case "mov"
        ' Quicktime Video
        l_TypeInfo = "video/quicktime"
    Case "mp4"
        ' MPEG-4 Video
        l_TypeInfo = "video/mp4"
    Case "mp4a"
        ' MPEG-4 Audio
        l_TypeInfo = "audio/mp4"
    Case "mpg"
        ' MPEG Video
        l_TypeInfo = "video/mpeg"
    Case "mpeg"
        ' MPEG Video
        l_TypeInfo = "video/mpeg"
    Case "mpga"
        ' MPEG Audio
        l_TypeInfo = "audio/mpeg"
    Case "mpp"
        ' Microsoft Project
        l_TypeInfo = "application/vnd.ms-project"
    Case "odb"
        ' OpenDocument Database
        l_TypeInfo = "application/vnd.oasis.opendocument.database"
    Case "odc"
        ' OpenDocument Chart
        l_TypeInfo = "application/vnd.oasis.opendocument.chart"
    Case "odf"
        ' OpenDocument Formula
        l_TypeInfo = "application/vnd.oasis.opendocument.formula"
    Case "odft"
        ' OpenDocument Formula Template
        l_TypeInfo = "application/vnd.oasis.opendocument.formula-template"
    Case "odg"
        ' OpenDocument Graphics
        l_TypeInfo = "application/vnd.oasis.opendocument.graphics"
    Case "odi"
        ' OpenDocument Image
        l_TypeInfo = "application/vnd.oasis.opendocument.image"
    Case "odm"
        ' OpenDocument Text Master
        l_TypeInfo = "application/vnd.oasis.opendocument.text-master"
    Case "odp"
        ' OpenDocument Presentation
        l_TypeInfo = "application/vnd.oasis.opendocument.presentation"
    Case "ods"
        ' OpenDocument Spreadsheet
        l_TypeInfo = "application/vnd.oasis.opendocument.spreadsheet"
    Case "odt"
        ' OpenDocument Text
        l_TypeInfo = "application/vnd.oasis.opendocument.text"
    Case "oga"
        ' Ogg Audio
        l_TypeInfo = "audio/ogg"
    Case "ogv"
        ' Ogg Video
        l_TypeInfo = "video/ogg"
    Case "ogx"
        ' Ogg
        l_TypeInfo = "application/ogg"
    Case "onetoc"
        ' Microsoft OneNote
        l_TypeInfo = "application/onenote"
    Case "otf"
        ' OpenType Font File
        l_TypeInfo = "application/x-font-otf"
    Case "pdf"
        ' Adobe Portable Document Format
        l_TypeInfo = "application/pdf"
    Case "png"
        ' Portable Network Graphics (PNG)
        l_TypeInfo = "image/png"
    Case "ppsx"
        ' Microsoft Office - OOXML - Presentation (Slideshow)
        l_TypeInfo = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
    Case "ppt"
        ' Microsoft PowerPoint
        l_TypeInfo = "application/vnd.ms-powerpoint"
    Case "pptx"
        ' Microsoft Office - OOXML - Presentation
        l_TypeInfo = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    Case "psd"
        ' Photoshop Document
        l_TypeInfo = "image/vnd.adobe.photoshop"
    Case "pub"
        ' Microsoft Publisher
        l_TypeInfo = "application/x-mspublisher"
    Case "qt"
        ' Quicktime Video
        l_TypeInfo = "video/quicktime"
    Case "ram"
        ' Real Audio Sound
        l_TypeInfo = "audio/x-pn-realaudio"
    Case "rar"
        ' RAR Archive
        l_TypeInfo = "application/x-rar-compressed"
    Case "rm"
        ' RealMedia
        l_TypeInfo = "application/vnd.rn-realmedia"
    Case "rss"
        ' RSS - Really Simple Syndication
        l_TypeInfo = "application/rss+xml"
    Case "svg"
        ' Scalable Vector Graphics (SVG)
        l_TypeInfo = "image/svg+xml"
    Case "swf"
        ' Adobe Flash
        l_TypeInfo = "application/x-shockwave-flash"
    Case "tar"
        ' Tar File (Tape Archive)
        l_TypeInfo = "application/x-tar"
    Case "tif"
        ' Tagged Image File Format
        l_TypeInfo = "image/tiff"
    Case "tiff"
        ' Tagged Image File Format
        l_TypeInfo = "image/tiff"
    Case "torrent"
        ' BitTorrent
        l_TypeInfo = "application/x-bittorrent"
    Case "tsv"
        ' Tab Seperated Values
        l_TypeInfo = "text/tab-separated-values"
    Case "ttf"
        ' TrueType Font
        l_TypeInfo = "application/x-font-ttf"
    Case "txt"
        ' Text File
        l_TypeInfo = "text/plain"
    Case "vcf"
        ' vCard
        l_TypeInfo = "text/x-vcard"
    Case "vcs"
        ' vCalendar
        l_TypeInfo = "text/x-vcalendar"
    Case "vsd"
        ' Microsoft Visio
        l_TypeInfo = "application/vnd.visio"
    Case "vsdx"
        ' Microsoft Visio 2013
        l_TypeInfo = "application/vnd.visio2013"
    Case "wav"
        ' Waveform Audio File Format (WAV)
        l_TypeInfo = "audio/x-wav"
    Case "weba"
        ' Open Web Media Project - Audio
        l_TypeInfo = "audio/webm"
    Case "webm"
        ' Open Web Media Project - Video
        l_TypeInfo = "video/webm"
    Case "webp"
        ' WebP Image
        l_TypeInfo = "image/webp"
    Case "wm"
        ' Microsoft Windows Media
        l_TypeInfo = "video/x-ms-wm"
    Case "wma"
        ' Microsoft Windows Media Audio
        l_TypeInfo = "audio/x-ms-wma"
    Case "wmf"
        ' Microsoft Windows Metafile
        l_TypeInfo = "application/x-msmetafile"
    Case "wmv"
        ' Microsoft Windows Media Video
        l_TypeInfo = "video/x-ms-wmv"
    Case "woff"
    Case "woff2"
        ' Web Open Font Format
        l_TypeInfo = "application/x-font-woff"
    Case "xhtml"
        ' XHTML - The Extensible HyperText Markup Language
        l_TypeInfo = "application/xhtml+xml"
    Case "xls"
        ' Microsoft Excel
        l_TypeInfo = "application/vnd.ms-excel"
    Case "xlsx"
        ' Microsoft Office OOXML - Spreadsheet
        l_TypeInfo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    Case "xml"
        ' XML - Extensible Markup Language
        l_TypeInfo = "application/xml"
    Case "xps"
        ' Microsoft XML Paper Specification
        l_TypeInfo = "application/vnd.ms-xpsdocument"
    Case "xslt"
        ' XML Transformations
        l_TypeInfo = "application/xslt+xml"
    Case "xul"
        ' XUL - XML User Interface Language
        l_TypeInfo = "application/vnd.mozilla.xul+xml"
    Case "zip"
        ' Zip Archive
        l_TypeInfo = "application/zip"
    Case "vbml"
        ' VBML - HTML files that are pre-processed by this
        l_TypeInfo = "text/html"
    Case Else
        ' Reached end of mime type info list by index. Return empty string to signal caller that we are done.
        l_TypeInfo = "application/octet-stream"
    End Select
    MapMimeType = l_TypeInfo
End Function
