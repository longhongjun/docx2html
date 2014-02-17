<?php
//unzip docx file
$file_name = 'example.docx';
if( !is_dir('./box')  ) {
    mkdir('./box', 0777);
    system("unzip -d ./box/ {$file_name} ");
}

//convert wmf to png
if( !is_dir('./image')  ) {
    mkdir('./image', 0777);
}

if ($handle = opendir('./box/word/media/')) {
    while (false !== ($entry = readdir($handle))) {
        if(!in_array($entry, array('.', '..'))) {
          $ar = explode('.', $entry);
          $cmd = "convert ./box/word/media/{$entry} ./image/{$ar[0]}.png";
          system($cmd);
        }
    }
    closedir($handle);
}

//relations of pic
$relation_str = file_get_contents('./box/word/_rels/document.xml.rels');
preg_match_all('/<Relationship Id="([^"]+)" Type="[^"]+" Target="media\/image(\d+).wmf"\/>/', $relation_str, $matches);
$relations = array_combine($matches[1], $matches[2]);

//portal
$content = file_get_contents('./box/word/document.xml');
$content = preg_replace_callback('/<v:imagedata r:id="([^"]+)" [^>]+>/', 'get_pic', $content);

echo $content;

//call back
function get_pic($matches) {
  global $relations;
  $id = $relations[$matches[1]];
  return '<img src="image/image' . $id .'.png" />';
}
