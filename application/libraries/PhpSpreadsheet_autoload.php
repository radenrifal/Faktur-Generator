<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');


class PhpSpreadsheet_autoload {
  public function __construct()
  {
    require_once APPPATH.'third_party/PhpOffice/autoload.php';
    require_once APPPATH.'third_party/Psr/autoload.php';        
  }
}