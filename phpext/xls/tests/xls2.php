<?
//	dl("xls.so");
	$wb=xls_open("Путь к файлу/test2.xls","KOI8-R");
	echo "<style type=\"text/css\">".xls_getcss($wb)."</style>";
	$list=0;
	echo "Charset:".xls_getcharset($wb)."<br>\n";
	echo "Sheets N:".xls_getsheetscount($wb)."<br>\n";
	$ws=xls_getworksheet($wb,$list);
	if (xls_parseworksheet($ws)) 
	{
		echo "Разборка листа \"".xls_getsheetname($wb,$list)."\" прошла успешно<br>\n";
		$fws=xls_fetch_worksheet($ws);
		echo "<table border=0 cellspacing=0 cellpadding=2>";
		for ($t=0;$t<=$fws->rows->lastrow;$t++)
		{
			echo "<tr>" ;
			for ($i=0;$i<=$fws->rows->lastcol;$i++) 
			{
				if ($fws->rows->row[$t]->cells->cell[$i]->ishiden==0)
				{
					echo "<td";
					if ($fws->rows->row[$t]->cells->cell[$i]->colspan) 
					echo " colspan=".$fws->rows->row[$t]->cells->cell[$i]->colspan;
					if ($fws->rows->row[$t]->cells->cell[$i]->rowspan) 
					echo " rowspan=".$fws->rows->row[$t]->cells->cell[$i]->rowspan;
					echo " class=xf".$fws->rows->row[$t]->cells->cell[$i]->xf;
					echo ">";
					if (isset($fws->rows->row[$t]->cells->cell[$i]->str))
						print  $fws->rows->row[$t]->cells->cell[$i]->str."&nbsp;";
					echo "</td>";
				}
			}
			echo "</tr>\n";
		}
		echo "</table>"
	}
?>