package LPR;
import java.io.*;
import java.net.*;
import java.nio.channels.SelectionKey;
import java.nio.channels.Selector;

import java.nio.channels.ServerSocketChannel;
import java.nio.channels.SocketChannel;
//import java.nio.CharBuffer;

public class Main {

    public static void main(String[] args) throws IOException {
        //=================================
        //Iniciar Valores
        //=================================
        int puerto_local;
        int puerto_remoto;
        InetAddress ip_remota;

        puerto_local  = Integer.parseInt(args[0]);
        puerto_remoto = Integer.parseInt(args[2]);
        ip_remota     = InetAddress.getByName(args[1]);

        System.out.println("Linux Port Redirect [tbrSoft]\n");
        System.out.println("Version 0.91.002\n\n");
        
        IniciarServicio(puerto_local, ip_remota, puerto_remoto);

    }

    private static void IniciarServicio(int puerto_local, InetAddress ip_remota, int puerto_remoto) throws IOException
    {



            //Sin esto, algunos caracteres se envian de manera incorrecta
            //No se por que... Lo copie y pegue
            String[] codePages = {// "8859_1", "CP1254", //NCR added these
                     "CP437", "CP737", "CP775", "CP850",
                     "CP852", "CP855", "CP857", "CP860",
                     "CP861", "CP862", "CP863", "CP864",
                     "CP865", "CP866", "CP869", "CP874",
                     "CP856", "CP858", "CP868", "CP870"
                    };

            String pl = puerto_local+"";
            String pr = puerto_remoto+"";
            String ipr = ip_remota+"";

            //=================================
            //Inicio el Servicio de Redireccion
            //=================================
            
            /*

             //1.
            //·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´` 
            //ssChannel abre el puerto al cual se conecta el Cliente
            ServerSocketChannel ssChannel = ServerSocketChannel.open();
            ssChannel.configureBlocking(false);
            ssChannel.socket().bind(new InetSocketAddress(puerto_local));
            //·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´` 


            System.out.println("Abrir Puerto "+pl);

            //skCliente es el manejador del Socket que esta conectado
            //al Cliente (una vez que se acepto la conexion)
            SocketChannel skCliente = ssChannel.accept();
             */


            ServerSocket SocketOpenCliente = new ServerSocket(puerto_local);
            //skCliente es el manejador del Socket que esta conectado
            Socket skCliente = SocketOpenCliente.accept();
            //Esperar a que se conecte cliente

            /*
            while (skCliente.)
            {
                skCliente = ssChannel.accept();
            }
            */
            skCliente.getReuseAddress();
            System.out.println("Conectado con cliente: "+skCliente.getRemoteSocketAddress()+"\n");
            //Ya se conecto el cliente!
            

            //READYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
            //2.
            //Conectarse con el Servidor
            System.out.println("Conectar con Servidor "+ipr+":"+pr);
            //skServer es el manejador del Socket que esta conectado
            //al servidor
            //·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`
            //(Esta linea es la que ordena la conexion al servidor mediante scChannel
            //InetSocketAddress scChannel = new InetSocketAddress(ip_remota, puerto_remoto);
            //ssChannel.socket().bind(new InetSocketAddress(puerto_local));
            //·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`·..·´`

            //InetSocketAddress skServer = new InetSocketAddress(ip_remota, puerto_remoto);


            //AQUI!!!!!!!!!!!!!!!!!
            Socket skServer = new Socket(ip_remota, puerto_remoto);

            //=================================
            //3. Iniciar Redireccionamiento
            //=================================
            
            //Estas son las variables que escuchan al Cliente
            //-------------------------------------------------
                
            //InputDataC es el manejador de los datos que entran
            BufferedReader InputDataC = new BufferedReader(new InputStreamReader(skCliente.getInputStream(), codePages[0]));
            //DataOut es el manejador de los datos que salen
            BufferedWriter DataOutC = new BufferedWriter(new OutputStreamWriter(skServer.getOutputStream(), codePages[0]));
            //en buffer_from_client tiene los datos que entran en el cliente
            char[] buffer_from_client = new char[1024];//1024=Leo buffers de 1Kb
            int len_buffer_cl=0;

                //Selector selectorCliente = Selector.open();
                //SelectionKey KeyCliente = skCliente.register(selectorCliente, SelectionKey.OP_ACCEPT);
             

            //Estas son las variables que escuchan al Servidor
            //-------------------------------------------------
            //InputData es el manejador de los datos que entran
            BufferedReader InputDataS = new BufferedReader(new InputStreamReader(skServer .getInputStream(), codePages[0]));
            //DataOut es el manejador de los datos que salen
            BufferedWriter DataOutS = new BufferedWriter(new OutputStreamWriter(skCliente.getOutputStream(), codePages[0]));
            //en buffer_from_client tiene los datos que entran en el cliente
            char[] buffer_from_server = new char[1024];//1024=Leo buffers de 1Kb
            int len_buffer_sr=0;
            


            //=================================
            //4. Redirecciono efectivamente
            //=================================
            while (true)
            {
                try
                {
                    //Entraron datos del Servidor al Cliente?
                    if(InputDataS.ready()==true)
                    //if((len_buffer_sr=InputDataS.read(buffer_from_server)) > 0)
                    {
                        len_buffer_sr=InputDataS.read(buffer_from_server);
                        //System.out.println(buffer_from_server);
                        
                        DataOutS.write(buffer_from_server, 0, len_buffer_sr);
                        DataOutS.flush();
                    }//fin iff

                    //Entraron datos del Cliente al Servidor?
                    if(InputDataC.ready()==true)
                    {
                        len_buffer_cl=InputDataC.read(buffer_from_client);
                        //System.out.println(buffer_from_client);

                        DataOutC.write(buffer_from_client, 0, len_buffer_cl);
                        DataOutC.flush();
                    }//fin iff
                } //fin try
                catch (Exception e)
                {
                    System.out.println(e.getMessage());
                } //fin catch
            }//fin while
        }// fin IniciarServicio()
 }