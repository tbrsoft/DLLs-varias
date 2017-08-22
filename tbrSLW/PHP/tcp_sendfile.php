<?php
    include_once 'GuardarArchivo.php';
?>

    <?php
    /*
     * Para enviar archivos mediante TCP
     * @author manuel
     */
    $save_file_path='';

    function recibir_archivo()
    {
        return(0);
    }

    class tcp_sendfile {

        //La carpeta donde se guarda la Licencia Activada
        function set_save_file_path($path_lic_activada)
        {
            $save_file_path=$path_lic_activada;
        }

        function get_buffer_archivo($path_archivo)
        {
            //rb=read binary
            $handle=fopen($path_archivo,'rb');

            $length=filesize($path_archivo);
            $buffer=fread($handle, $length);
            fread($handle, $length);
            fclose($handle);

            return $buffer;
        }

        /*
         * Valores que devuelve enviar_archivo
         * 0 - Sin Creditos
         * 1 - Licencia Activada
         */
        function enviar_archivo($server_ip,$puerto,$nombe_archivo,$buffer)
        {
            $Guardador = new GuardarArchivo();
            //Crear un socket
            $socket = socket_create(AF_INET, SOCK_STREAM, SOL_TCP);
            //$fp = fsockopen($server_ip, $puerto, $errno, $errstr, 10);
            if ($socket === false) {
                echo "socket_create() falló: " . socket_strerror(socket_last_error()) . "\n";
            } else {
                //echo "Socket Ok";
            }


            //Conectarse al servidor
            echo "Intentando conectar a '$server_ip':'$puerto'...";
            $result = socket_connect($socket, $server_ip, $puerto);
            if ($result === false) {
                echo "socket_connect() falló.\n ($result) " . socket_strerror(socket_last_error($socket)) . "\n";
            } else {

                echo '<BR><BR>Conectado.<BR><BR>';
                //====================================================
                $separador='//';
                $codigo_archivo_recibido='002';
                $codigo_archivo_recibido.=$separador;
                $archivo_recibido='El Archivo fue recibido, el servidor esta trabajando
                    espere un momento';
                //====================================================

                //si se logro la conexion, se envia:
                //001//IdUsuario//NombreArchivo//LargoArchivo::BufferArchivoTotal

                $buffer_tcp='001';
                $buffer_tcp.=$separador;
                $buffer_tcp.='145';//IdUsuario
                $buffer_tcp.=$separador;
                $buffer_tcp.=$nombe_archivo;
                $buffer_tcp.=$separador;
                $buffer_tcp.=strlen($buffer);
                $buffer_tcp.='::';
                $buffer_tcp.=$buffer;

                $estoyConectado=true;

                //envio el paquete de datos (La Licencia)
                
                socket_send($socket, $buffer_tcp, strlen($buffer_tcp), MSG_DONTROUTE);
                //fwrite($fp, $buffer_tcp);

                $final_respuesta='';

                while ($estoyConectado==true) {
                    //$respuesta=fgets($fp);
                    //$respuesta='';

                    //$recvs=socket_recv($socket, $respuesta, 1024, 0);
                    $recvs=socket_recv($socket, $respuesta, 8192, 16);
                    if($recvs > 0)
                    {
                        //Leo los 3 primeros caracteres de lo que entro
                        $codigo=substr($respuesta, 0, 3);

                        //fwrite($fp, $codigo_archivo_recibido);

                        switch ($codigo)
                        {
                            case '001':
                                //recibir_archivo();
                                $final_respuesta.='La Licencia Activada se recibio correctamente<BR>';
                                //Le digo al server que recibi la licencia
                                socket_send($socket, $codigo_archivo_recibido, strlen($codigo_archivo_recibido), MSG_DONTROUTE);
                                $Guardador->GuardarArchivo($respuesta);

                                //Como termine me desconecto (Salgo de while)
                                $estoyConectado=false;
                                break;

                            case '002':
                                $final_respuesta.='Servidor Recibio Archivo<BR>';
                                //$final_respuesta.=$respuesta;
                                break;
                            case '099':
                                print('<BR>Desconectarse<BR>');
                                $estoyConectado=false;
                                break;
                        }
                    }//fin if recvs
                }//fin while
                echo($final_respuesta);
                echo('<BR> FIN.xxx');
                //socket_close($fp);
                socket_shutdown($fp);

                return(0);
            }
        }
    }
?>
