<?php

spl_autoload_register(function ($class_name) {
  $path_to_file = str_replace("Psr", "", dirname(__FILE__)) . str_replace("\\", DIRECTORY_SEPARATOR, $class_name) . ".php";
	try{
		include $path_to_file;
	}
	catch(Exception $e){
		echo "error during loading";
	}  
});