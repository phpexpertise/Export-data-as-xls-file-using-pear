<?php
/**
 * Author : phpexpertise <info@phpexpertise.com>
 * Description : Connect to the database and config settings.
 */
define('HOSTNAME',"localhost");
define('USERNAME',"root");
define('PASSWORD',"");
define('DATABASE',"test");
define('BASEPATH',"<Your Localhost>");

try{
    $db_connect    =    new PDO("mysql:host=".HOSTNAME.";dbname=".DATABASE,USERNAME,PASSWORD);    
}
catch(PDOException $e){
    $e->getMessage();
}

spl_autoload_register(function($class_name){
    include_once $class_name.'.php';
});
?>