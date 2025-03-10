xquery version "3.0";
declare namespace output = "http://www.w3.org/2010/xslt-xquery-serialization"; 
declare option output:method "html"; 
declare option output:indent "yes"; 
declare option output:omit-xml-declaration "yes";

<html>
    <head><title>List of assets</title></head>
    <body>
        <table>
            <tr>
                <th>Unit</th> <th>Session</th> <th>Type</th> <th>Image</th> <th>ImageSrc</th> <th>Caption</th> <th>Rights</th>
            </tr>
            {
                for $mm in .//MediaContent
                let $src := $mm/@src
                let $type := $mm/@type
                let $id := $mm/@id
                let $week := $mm/ancestor::Unit/UnitTitle/string()
                let $session:= $mm/ancestor::Session/Title/string()
                let $caption := $mm/Caption/string()
                let $rights := $mm/SourceReference/string()
                return
                    <tr>
                        <td>{$week}</td>
                        <td>{$session}</td>
                        <td>{string($type)}</td>
                        <td>{string($id)}</td>
                        <td>{string($src)}/</td>
                        <td>{$caption}</td>
                        <td>{$rights}</td>
                    </tr>
            }
            {
                for $fig in .//Figure
                let $src := $fig/Image/@src
                let $week := $fig/ancestor::Unit/UnitTitle/string()
                let $session:= $fig/ancestor::Session/Title/string()
                let $caption := $fig/Caption/string()
                let $rights := $fig/SourceReference/string()
                return
                    <tr>
                        <td>{$week}</td>
                        <td>{$session}</td>
                        <td>{"Image"}</td>
                        <td><img src="{$src}" width="100"/></td>
                        <td>{string($src)}/</td>
                        <td>{$caption}</td>
                        <td>{$rights}</td>
                    </tr>
            }
        </table>
    </body>
</html>
