<?php
/* 
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of GuardarArchivo
 *
 * @author manuel
 */
class GuardarArchivo {
    function GuardarArchivo($buffer)
    {
        //cargo los datos desde el byte 1 hasta ::
        $aux=substr($buffer,1,strpos($buffer, "::"));
        
        //$usuario
        $paht='\Archivos\\'+$usuario+'\\'+$nombre_archivo;
        
    }

}
?>
