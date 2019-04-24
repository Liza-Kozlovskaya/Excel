<html>
    <head></head>
    <body>
        <!-- 1 -->
        <form action="table.php" method="POST" enctype="multipart/form-data">
        <input type="file" name="upload">
        <input type="submit" value="Обработать">
        </form>
    </body>
</html>
<?php
    require_once 'A:\home\test.com\www\PHPExcel-1.8\Classes\PHPExcel.php';
    require_once 'A:\home\test.com\www\PHPExcel-1.8\Classes\PHPExcel\IOFactory.php';
    $file = $_FILES['upload']['name'];
    //$file = 'A:\home\test.com\www\test.xlsx';
    
    function ProcessingFile($file)
    {
        $fileexcel = PHPExcel_IOFactory::load($file);
        foreach($fileexcel ->getWorksheetIterator() as $worksheet)
        {
            $lists[] = $worksheet->toArray();
        }
        SerchCountEmptyCol($lists);
    }


    function WorksheetUSBCharger($file, $i)
    {
        $fileexcel = PHPExcel_IOFactory::load($file);
        if($i==4)
        {
            $worksheet = $fileexcel->getSheet(2);
            $ls[] = $worksheet->toArray();
            ShowTable($ls);
        }
        else if($i==5)
        {
            $worksheet = $fileexcel->getSheet(1);
            $ls[] = $worksheet->toArray();
            ShowTable($ls);
        }

    }

    function ShowTable($lists)
    {
        foreach($lists as $list)
        {
            //перебор строк
            echo '<table border="1">';
            foreach($list as $row)
            {
                //перебор столбцов
                echo '<tr>';
                foreach($row as $col)
                {
                    echo '<td>'.$col.'</td>';
                }
                echo '</tr>';
            }
            echo '</table>';
        }
    }

    function CountList($file)
    {
        $objPHPExcel = PHPExcel_IOFactory::load($file);
        foreach($objPHPExcel->getAllSheets() as $sheet)
        {
            $sheets[] = $sheet->getTitle();
        }
        ShowList($sheets);
    }

    function ShowList($sheets)
    {
        $num =1;
        echo '<br><table border="1">';
        foreach($sheets as $sh)
        {   
            echo '<tr><td>' . $num++ . '</td><td>'.$sh.'</td></tr>';
        }
        echo '</table>';
    }

    echo "<b>Все данные в файле</b>"; 
    ProcessingFile($file); //2
    echo "<br> <b>Названия всех листов в файле</b>";
    CountList($file); //3
    echo "<br> <b>Данные из листа USBCharger</b>";
    WorksheetUSBCharger($file, $i=4); //4
    echo "<br> <b>Данные из листа Holder</b>";
    WorksheetUSBCharger($file, $i=5); //5
?>
