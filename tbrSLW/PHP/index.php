<html>
    <form enctype="multipart/form-data" action="enviararchivo.php" method="post">
        <input type="hidden" name="MAX_FILE_SIZE" value="50000" />
        Archivo de Licencia <BR> 
        <input type="file" name="licen" size="40">
        <BR>
        <input type="submit" value="Enviar">

    </form>
</html>
<?php

    /*
    //Tamaño maximo del archivo = 50Kb
    echo('<input type="hidden" name="MAX_FILE_SIZE" value="50000" />>');
    echo('<input type="file" name="Archivo de Licencia <BR>" size="140">');
    echo('<BR>');
    $path_archivo=$_FILES['uploadedfile'];
    echo('<input type="submit" value="Enviar" onclick="enviararchivo.php?archivo=$path_archivo">');
    */
?>
