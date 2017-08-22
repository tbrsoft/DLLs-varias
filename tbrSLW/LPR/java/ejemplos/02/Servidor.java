// Configurar un servidor que reciba una conexi�n de un cliente, env�e
// una cadena al cliente y cierre la conexi�n.
import java.io.*;
import java.net.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;

public class Servidor extends JFrame {
   private JTextField campoIntroducir;
   private JTextArea areaPantalla;
   private ObjectOutputStream salida;
   private ObjectInputStream entrada;
   private ServerSocket servidor;
   private Socket conexion;
   private int contador = 1;

   // configurar GUI
   public Servidor()
   {
      super( "Servidor" );

      Container contenedor = getContentPane();

      // crear campoIntroducir y registrar componente de escucha
      campoIntroducir = new JTextField();
      campoIntroducir.setEditable( false );
      campoIntroducir.addActionListener(
         new ActionListener() {

            // enviar mensaje al cliente
            public void actionPerformed( ActionEvent evento )
            {
               enviarDatos( evento.getActionCommand() );
               campoIntroducir.setText( "" );
            }
         }  
      ); 

      contenedor.add( campoIntroducir, BorderLayout.NORTH );

      // crear areaPantalla
      areaPantalla = new JTextArea();
      contenedor.add( new JScrollPane( areaPantalla ), 
         BorderLayout.CENTER );

      setSize( 300, 150 );
      setVisible( true );

   } // fin del constructor de Servidor

   // configurar y ejecutar el servidor 
   public void ejecutarServidor()
   {
      // configurar servidor para que reciba conexiones; procesar las conexiones
      try {

         // Paso 1: crear un objeto ServerSocket.
         servidor = new ServerSocket( 12345, 100 );

         while ( true ) {

            try {
               esperarConexion(); // Paso 2: esperar una conexi�n.
               obtenerFlujos();        // Paso 3: obtener flujos de entrada y salida.
               procesarConexion(); // Paso 4: procesar la conexi�n.
            }

            // procesar excepci�n EOFException cuando el cliente cierre la conexi�n 
            catch ( EOFException excepcionEOF ) {
               System.err.println( "El servidor termin� la conexi�n" );
            }

            finally {
               cerrarConexion();   // Paso 5: cerrar la conexi�n.
               ++contador;
            }

         } // fin de instrucci�n while

      } // fin del bloque try

      // procesar problemas con E/S
      catch ( IOException excepcionES ) {
         excepcionES.printStackTrace();
      }

   } // fin del m�todo ejecutarServidor

   // esperar que la conexi�n llegue, despu�s mostrar informaci�n de la conexi�n
   private void esperarConexion() throws IOException
   {
      mostrarMensaje( "Esperando una conexi�n\n" );
      conexion = servidor.accept(); // permitir al servidor aceptar la conexi�n            
      mostrarMensaje( "Conexi�n " + contador + " recibida de: " +
         conexion.getInetAddress().getHostName() );
   }

   // obtener flujos para enviar y recibir datos
   private void obtenerFlujos() throws IOException
   {
      // establecer flujo de salida para los objetos
      salida = new ObjectOutputStream( conexion.getOutputStream() );
      salida.flush(); // vaciar b�fer de salida para enviar informaci�n de encabezado

      // establecer flujo de entrada para los objetos
      entrada = new ObjectInputStream( conexion.getInputStream() );

      mostrarMensaje( "\nSe recibieron los flujos de E/S\n" );
   }

   // procesar la conexi�n con el cliente
   private void procesarConexion() throws IOException
   {
      // enviar mensaje de conexi�n exitosa al cliente
      String mensaje = "Conexi�n exitosa";
      enviarDatos( mensaje );

      // habilitar campoIntroducir para que el usuario del servidor pueda enviar mensajes
      establecerCampoTextoEditable( true );

      do { // procesar los mensajes enviados por el cliente

         // leer el mensaje y mostrarlo en pantalla
         try {
            mensaje = ( String ) entrada.readObject();
            mostrarMensaje( "\n" + mensaje );
         }

         // atrapar problemas que pueden ocurrir al tratar de leer del cliente
         catch ( ClassNotFoundException excepcionClaseNoEncontrada ) {
            mostrarMensaje( "\nSe recibi� un tipo de objeto desconocido" );
         }

      } while ( !mensaje.equals( "CLIENTE>>> TERMINAR" ) );

   } // fin del m�todo procesarConexion

   // cerrar flujos y socket
   private void cerrarConexion() 
   {
      mostrarMensaje( "\nFinalizando la conexi�n\n" );
      establecerCampoTextoEditable( false ); // deshabilitar campoIntroducir

      try {
         salida.close();
         entrada.close();
         conexion.close();
      }
      catch( IOException excepcionES ) {
         excepcionES.printStackTrace();
      }
   }

   // enviar mensaje al cliente
   private void enviarDatos( String mensaje )
   {
      // enviar objeto al cliente
      try {
         salida.writeObject( "SERVIDOR>>> " + mensaje );
         salida.flush();
         mostrarMensaje( "\nSERVIDOR>>> " + mensaje );
      }

      // procesar problemas que pueden ocurrir al enviar el objeto
      catch ( IOException excepcionES ) {
         areaPantalla.append( "\nError al escribir objeto" );
      }
   }

   // m�todo utilitario que es llamado desde otros subprocesos para manipular a
   // areaPantalla en el subproceso despachador de eventos
   private void mostrarMensaje( final String mensajeAMostrar )
   {
      // mostrar mensaje del subproceso de ejecuci�n despachador de eventos
      SwingUtilities.invokeLater(
         new Runnable() {  // clase interna para asegurar que la GUI se actualice apropiadamente

            public void run() // actualiza areaPantalla
            {
               areaPantalla.append( mensajeAMostrar );
               areaPantalla.setCaretPosition( 
                  areaPantalla.getText().length() );
            }

         }  // fin de la clase interna

      ); // fin de la llamada a SwingUtilities.invokeLater
   }

   // m�todo utilitario que es llamado desde otros subprocesos para manipular a 
   // campoIntroducir en el subproceso despachador de eventos
   private void establecerCampoTextoEditable( final boolean editable )
   {
      // mostrar mensaje del subproceso de ejecuci�n despachador de eventos
      SwingUtilities.invokeLater(
         new Runnable() {  // clase interna para asegurar que la GUI se actualice apropiadamente

            public void run()  // establece la capacidad de modificar a campoIntroducir
            {
               campoIntroducir.setEditable( editable );
            }

         }  // fin de la clase interna

      ); // fin de la llamada a SwingUtilities.invokeLater
   }

   public static void main( String args[] )
   {
      Servidor aplicacion = new Servidor();
      aplicacion.setDefaultCloseOperation( JFrame.EXIT_ON_CLOSE );
      aplicacion.ejecutarServidor();
   }

}  // fin de la clase Servidor