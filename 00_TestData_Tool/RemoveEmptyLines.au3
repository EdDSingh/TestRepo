#include <FileConstants.au3>
Local $vfilePath_1 = $CmdLine[1]
$handle = FileOpen($vfilePath_1) ; open file (read mode)
$sContent = FileRead($handle) ; read content
FileClose($handle)

Do ; remove double carriage-returns (white lines)
    $sContent = StringReplace($sContent, @CRLF & @CRLF, @CRLF)
Until Not @extended

$handle = FileOpen($vfilePath_1, $FO_OVERWRITE) ; write result in a new file
FileWrite($handle, $sContent)
FileFlush($handle)
FileClose($handle)