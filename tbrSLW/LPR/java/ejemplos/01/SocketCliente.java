/*
 * Javier Abellán. 27 de noviembre de 2003
 *
 * SocketCliente.java
 *
 */

import java.net.*;
import java.io.*;

/**
 * Clase que crea un socket cliente, establece la conexión y lee los datos
 * del servidor, escribiéndolos en pantalla.
 */
public class SocketCliente
 {
     /** Programa principal, crea el socket cliente */
     public static void main (String [] args)
     {
         new SocketCliente();
     }
     
     /**
      * Crea el socket cliente y lee los datos
      */
     public SocketCliente()
     {
         try
         {
             /* Se crea el socket cliente */
             Socket socket = new Socket ("localhost", 35557);
             System.out.println ("conectado");
             
             /* Se obtiene un stream de lectura para leer tipos simples de java */
             DataInputStream buffer = new DataInputStream (socket.getInputStream());
             
             /**Se lee un entero y un String que nos envía el servidor, 
              escribiendo el resultado en pantalla */
             System.out.println("Recibido " + buffer.readInt());
             System.out.println ("Recibido " + buffer.readUTF());
             
             /* Se obtiene un stream de lectura para leer objetos */
             ObjectInputStream bufferObjetos =
                new ObjectInputStream (socket.getInputStream());
             
             /* Se lee un Datosocket que nos envía el servidor y se muestra 
              * en pantalla */
             DatoSocket dato = (DatoSocket)bufferObjetos.readObject();
             System.out.println ("Recibido " + dato.toString());
         }
         catch (Exception e)
         {
             e.printStackTrace();
         }
     }
}
