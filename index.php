<?php
/**
 * Author : phpexpertise <info@phpexpertise.com>
 * Load config file and excel library file.
 */
require_once 'config.php';

$xls_pear_obj    =    new ExcelPear($db_connect);

// SET TAB NAME AND FILE NAME
$xls_pear_obj->setTitle("users");

// EMPTY TEMPLATE
$xls_pear_obj->export_users();

?>