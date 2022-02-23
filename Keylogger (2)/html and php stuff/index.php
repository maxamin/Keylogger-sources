<?php
$msg = $_GET['test'];
$logfile= 'log.html';
$fp = fopen($logfile, "a"); 
fwrite($fp, $msg);
fclose($fp); 
?>