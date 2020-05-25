<?php
if ($_FILES["file"]["error"] > 0)
{
    echo "错误：" . $_FILES["file"]["error"] . "<br>";
}
else
{
    if (file_exists("sheet.xlsx")){
        unlink("sheet.xlsx");
    }
        
    if(move_uploaded_file($_FILES["file"]["tmp_name"], "sheet.xlsx")){
        header("Location:outputData.php");
    }
    else{
        echo("上传失败！");
    }
}
?>