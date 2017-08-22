<?php
    $server_ip = '200.81.207.65';
    $puerto=21;
    $timeout=3;
    //$sk=fsockopen($server_ip,$puerto,$errnum,$errstr,$timeout) ;

    $socket = socket_create(AF_INET, SOCK_STREAM, SOL_TCP);
    if ($socket === false) {
        echo "socket_create() falló: " . socket_strerror(socket_last_error()) . "\n";
    } else {
        //echo "Socket Ok";
    }

    echo "Intentando conectar a '$server_ip':'$puerto'...";
    $result = socket_connect($socket, $server_ip, $puerto);
    if ($result === false) {
        echo "socket_connect() falló.\n ($result) " . socket_strerror(socket_last_error($socket)) . "\n";
    } else {
        echo "Conectado a '$server_ip':'$puerto'...";
        flush();

        socket_send($socket, 'Hola Che!', 1024, MSG_DONTROUTE);

        $dati="" ;
        $dati2='Immer die Anderen';
        //socket_get_status($socket);
        //while (!feof($socket)) {
        for($i=0;$i<2;$i++) {
            //$dati = fgets ($sk, 1024);
            socket_recv($socket, $dati, 1024, 16);
            echo('<BR><BR>');
            echo($dati);
            echo('<BR><BR>****************************************************');

            //=============================
            $codigo=substr($dati, 0, 3);


            switch ($codigo)
            {
                case '001':
                    $dati2.='Bien Chori';
                    break;
                case '002':
                    $dati2.='Bien2Chori';
                    break;
            }//fin switch
            //=============================
        }//fin while

    }//fin else
    //fclose($sk) ;
    echo($dati2) ;

?>
