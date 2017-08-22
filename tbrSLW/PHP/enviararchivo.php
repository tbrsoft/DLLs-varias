<?php
    include_once 'tcp_sendfile.php';
?>

<?php

    echo('Activando Licencia...');
    flush();
    
    //$archivo=$_REQUEST[licen];
    $archivo=$_FILES['licen']['tmp_name'];
    $archivo_nombre=$_FILES['licen']['name'];

    $server_ip = '200.81.207.90';
    $puerto=21;
    $tcp_f = new tcp_sendfile();

    echo('Cargando ');
    echo($archivo_nombre);
    echo('<BR>');
    
    $buffer = $tcp_f->get_buffer_archivo($archivo);

    echo('Enviando ');
    echo($archivo_nombre);
    echo('<BR>');
    for($i=0;$i<10;$i++)
 
    /*
    ob_flush();
    flush();
    ob_end_flush();
    */

    $tcp_f->enviar_archivo($server_ip, $puerto, $archivo_nombre, $buffer);
    exit();
?>
