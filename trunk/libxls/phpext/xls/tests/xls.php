<?
	$wb=xls_open("Path to file/test2.xls","KOI8-R");
	$styles=xls_getcss($wb);
	$charset=xls_getcharset($wb);
	$sheets=xls_getsheetscount($wb);
	$list=0;

	echo "<pre>[style type=\"text/css\"]\n".$styles."[/style]</pre>";
	echo "Charset:".$charset."<br>";
	echo "Sheets:".$sheets."<br>";
        echo "Page \"".xls_getsheetname($wb,$list)."\" parsing...<br>";
	$ws=xls_getworksheet($wb,$list);
	if (xls_parseworksheet($ws)) 
	{
		echo "Ok<br>";
		$fws=xls_fetch_worksheet($ws);
		echo "<pre>";
		print_r($fws);
		echo "</pre>";
	}

?>