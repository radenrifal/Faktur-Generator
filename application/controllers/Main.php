<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Main extends CI_Controller {

  public function index()
  {
    $this->load->view('main');
  }

  public function uploadFile() 
  {
    $arrResult = array('success' => 1, 'message' => 'data telah berhasil d upload');

    $tmp_name = $_FILES["file"]["tmp_name"];
    $fileName = $_FILES["file"]["name"];

    $targetUpload = FCPATH.'assets/files-process/'.$fileName;

    if (!move_uploaded_file($tmp_name, $targetUpload)) {
      $arrResult = array('success' => 0, 'message' => 'gagal melakukan upload file');
    }

    $arrResult['file_name'] = $fileName;

    echo json_encode($arrResult);
    exit();
  }

  public function readFile() 
  {
    $this->load->library('PhpSpreadsheet_autoload');

    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true);
    $spreadsheet = $reader->load(FCPATH . 'assets/files-process/'.$this->input->post('file_name'));
    $sheet = $spreadsheet->getSheet($spreadsheet->getFirstSheetIndex());
    $data = $sheet->toArray();

    // output the data to the console, so you can see what there is.

    /*
    select 'FK' as code, '01' as jenis_transaksi, 0 as fp_pengganti, invoice_tax_number as nomor_faktur,fod.id_fac_invoice_out,extract(month from invoice_date) as masa_pajak,
                extract(year from invoice_date) as tahun_pajak, invoice_date as tanggal_faktur, s.npwp, customer_name as nama, s.address as alamat_lengkap,
                jumlah_dpp, jumlah_ppn, 0 as jumlah_ppnbm, '' as id_keterangan_tambahan, 0 as fg_uang_muka, 0 as uang_muka_dpp, 0 as uang_muka_ppn, 0 as uang_muka_ppnbm, '' as referensi
    */

    $arrMappingHeader = array(
      'id_fac_invoice_out' => 0,
      'nomor_faktur_num' => 4,
      'nomor_faktur_code' => 15,
      'invoice_date' => 5,
      'npwp' => 59,
      'nama' => 7,
      'alamat_lengkap' => 64,
      'jumlah_dpp' => 37,
      'jumlah_ppn' => 35,
    );


/*
SELECT    'OF' AS code, inv.*, case when dod.id is null then inv.item_name else ip.product_code end as kode_objek,
                        case when dod.id is null then inv.item_name else ip.product_name end as nama, 0 as tarif_ppnbm, 0 as ppnbm, dod.id_so_detail, 
                        case when dod.id is null then 1 else (dod.qty-coalesce(retur.qty_retur,0)) end as jumlah_barang,        
                        CASE WHEN i.invoice_date >= DATE '2022-04-01' 
                        THEN                 
                          case when dod.id is null then inv.home_base_amount/1.11 else sod.unit_price/1.11 end                        
                        ELSE 
                          case when dod.id is null then inv.home_base_amount/1.1 else sod.unit_price/1.1 end                        
                        END AS harga_satuan,
                        CASE WHEN i.invoice_date >= DATE '2022-04-01'
                        THEN
                          sod.discount_amount/1.11 * (case when dod.id is null then 1 else dod.qty end)
                        ELSE
                          sod.discount_amount/1.1 * (case when dod.id is null then 1 else dod.qty end)
                        END AS diskon,
                        CASE WHEN i.invoice_date >= DATE '2022-04-01'
                        THEN
                          (sod.unit_price/1.11) * (case when dod.id is null then 1 else (dod.qty-coalesce(retur.qty_retur,0)) end)
                        ELSE
                          (sod.unit_price/1.1) * (case when dod.id is null then 1 else (dod.qty-coalesce(retur.qty_retur,0)) end)
                        END AS harga_total,
                        CASE WHEN i.invoice_date >= DATE '2022-04-01'
                        THEN
                          sod.total_price/1.11
                        ELSE
                          sod.total_price/1.1
                        END AS dpp,     
                        CASE WHEN i.invoice_date >= DATE '2022-04-01'
                        THEN
                          (sod.total_price/1.11) * 11/100
                        ELSE
                          (sod.total_price/1.1) * 10/100
                        END AS  ppn,                             
                        it.home_base_amount AS inv_ppn
*/
    $arrMappingDetail = array(
      'kode_objek' => 104,
      'nama' => 109,
      'jumlah_barang' => 126,
      'harga_satuan' => 132, 
      'diskon' => 123,
      'ppn' => 127,
    );


    $arrDataResult = array();
    $currentInvoiceNum = '';
    foreach($data AS $idx => $row) {
      if ($idx < 2) continue;

      if ($row[0] != '') {
        $totalDpp = 0;
        $totalPpn = 0;

        $currentInvoiceNum = $row[0];
        $arrDataResult[$currentInvoiceNum] = array(
          'code' => 'FK',
          'jenis_transaksi' => '01',
          'fp_pengganti' => 0,
          'jumlah_ppnbm' => 0,
          'id_keterangan_tambahan' => '',
          'fg_uang_muka' => 0,
          'uang_muka_dpp' => 0,
          'uang_muka_ppn' => 0,
          'uang_muka_ppnbm' => 0,
          'referensi' => 0,
          'nomor_faktur_num' => '',
          'nomor_faktur_code' => '',
          'detail' => array()
        );

        foreach ($arrMappingHeader AS $idxName => $colNum) {
          $val = (empty($row[$colNum])) ? '' : $row[$colNum];
          $arrDataResult[$currentInvoiceNum][$idxName] = $val;
        }

        $arrDataResult[$currentInvoiceNum]['nomor_faktur'] = $arrDataResult[$currentInvoiceNum]['nomor_faktur_num'];
      }


      $arrTmpDetail = array(
        'code' => 'OF',
        'tarif_ppnbm' => 0,
        'ppnbm' => 0,
      );

      foreach ($arrMappingDetail AS $idxName => $colNum) {
        $val = (empty($row[$colNum])) ? '' : $row[$colNum];
        $arrTmpDetail[$idxName] = $val;
      }

      $arrTmpDetail['harga_total'] = round(floatval($arrTmpDetail['harga_satuan'])) * floatval($arrTmpDetail['jumlah_barang']);
      $arrTmpDetail['harga_dpp'] = floatval($arrTmpDetail['harga_total']) - round(floatval($arrTmpDetail['diskon']));

      $totalDpp += floatval($arrTmpDetail['harga_dpp']);
      $totalPpn += round(floatval($arrTmpDetail['ppn']));

      $arrDataResult[$currentInvoiceNum]['jumlah_dpp'] = $totalDpp;
      $arrDataResult[$currentInvoiceNum]['jumlah_ppn'] = $totalPpn;

      array_push($arrDataResult[$currentInvoiceNum]['detail'], $arrTmpDetail);
    }

    //$fileResult = $this->writeFileExcel($arrDataResult);
    $fileNameCsv = 'e_faktur_'.date('YmdHis').'.csv';
    $fileResult = $this->writeFileCSV($arrDataResult, $fileNameCsv);
    echo json_encode(array('success' => 1, 'filepath' => $fileResult['filepath'], 'filename' => $fileResult['filename']));
    // die(print_r($data[4], true)); 
    exit();
  }

  public function writeFileExcel($arrDataResult) 
  {
    $templateFile   = FCPATH.'assets/e_faktur.xlsx';
    $outputFileName = FCPATH.'assets/files-output/e_faktur.xlsx';
    $outputFilePath = base_url().'assets/files-output/e_faktur.xlsx';

    $reader         = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet    = $reader->load($templateFile);
    
    $sheet          = $spreadsheet->getActiveSheet();
    $currentRow     = 4;
    if(!empty($arrDataResult))
    {
      foreach($arrDataResult AS $idx => $row)
      {
        $sheet->setCellValue('A'.$currentRow, $row['code']);
        $sheet->getStyle('A'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->setCellValueExplicit('B'.$currentRow, $row['jenis_transaksi'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet->setCellValueExplicit('C'.$currentRow, $row['fp_pengganti'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet->getStyle('C'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
        $sheet->setCellValueExplicit('D'.$currentRow, "", \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet->getStyle('D'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        
        //proses penentuan masa dan tahun pajak ambil dari bulan dan tahun dari invoice_date
        $tglFaktur  = "";
        $masaPajak  = "";
        $tahunPajak = "";
        if(!empty($row['invoice_date']))
        {
          $tglFaktur  = date("d/m/Y", strtotime($row['invoice_date']));
          $masaPajak  = date("m", strtotime($row['invoice_date']));
          $tahunPajak = date("Y", strtotime($row['invoice_date']));
        }

        $sheet->setCellValue('E'.$currentRow, intval($masaPajak));
        $sheet->setCellValue('F'.$currentRow, intval($tahunPajak));
        $sheet->setCellValue('G'.$currentRow, $tglFaktur);
        $sheet->getStyle('G'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
        
        //hilangkan tanda baca terlebih dahulu
        $npwp = str_replace('.', '', $row['npwp']);
        $npwp = str_replace('-', '', $npwp);
        $sheet->setCellValueExplicit('H'.$currentRow, $npwp, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet->getStyle('H'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('I'.$currentRow, strtoupper($row['nama']));
        $sheet->setCellValue('J'.$currentRow, strtoupper($row['alamat_lengkap']));
        $sheet->setCellValueExplicit('K'.$currentRow, round($row['jumlah_dpp']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('K'.$currentRow)->getNumberFormat()->setFormatCode('#,##0'); 
        $sheet->setCellValueExplicit('L'.$currentRow, round($row['jumlah_ppn']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('L'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValueExplicit('M'.$currentRow, round($row['jumlah_ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('M'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValue('N'.$currentRow, $row['id_keterangan_tambahan']);
        $sheet->setCellValueExplicit('O'.$currentRow, round($row['fg_uang_muka']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('O'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValueExplicit('P'.$currentRow, round($row['uang_muka_dpp']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('P'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValueExplicit('Q'.$currentRow, round($row['uang_muka_ppn']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('Q'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValueExplicit('R'.$currentRow, round($row['uang_muka_ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->getStyle('R'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');
        $sheet->setCellValue('S'.$currentRow, $row['referensi']);
        
        //jika detailnya ada maka catat detail
        if(!empty($row['detail']))
        {
          $arrDetail = $row['detail'];

          foreach($arrDetail AS $rowDetail)
          {
            ++$currentRow;
            $sheet->setCellValue('A'.$currentRow, $rowDetail['code']);
            $sheet->getStyle('A'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $sheet->setCellValueExplicit('B'.$currentRow, $rowDetail['kode_objek'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValueExplicit('C'.$currentRow, $rowDetail['nama'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);

            // $sheet->setCellValueExplicit('D'.$currentRow, round(floatval($rowDetail['harga_satuan'])), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('D'.$currentRow, floatval($rowDetail['harga_satuan']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('D'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('D'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            //$sheet->setCellValueExplicit('E'.$currentRow, round($rowDetail['jumlah_barang']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('E'.$currentRow, $rowDetail['jumlah_barang'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('E'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('E'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

            // $sheet->setCellValueExplicit('F'.$currentRow, round($rowDetail['harga_total']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('F'.$currentRow, $rowDetail['harga_total'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('F'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('F'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            // $sheet->setCellValueExplicit('G'.$currentRow, round(floatval($rowDetail['diskon'])), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('G'.$currentRow, floatval($rowDetail['diskon']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('G'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('G'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            $sheet->setCellValueExplicit('H'.$currentRow, floatval($rowDetail['harga_dpp']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('H'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('H'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            // $sheet->setCellValueExplicit('I'.$currentRow, round($rowDetail['ppn']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('I'.$currentRow, floatval($rowDetail['ppn']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('I'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('I'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            // $sheet->setCellValueExplicit('J'.$currentRow, round($rowDetail['tarif_ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('J'.$currentRow, floatval($rowDetail['tarif_ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('J'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('J'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            
            // $sheet->setCellValueExplicit('K'.$currentRow, round($rowDetail['ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('K'.$currentRow, floatval($rowDetail['ppnbm']), \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->getStyle('K'.$currentRow)->getNumberFormat()->setFormatCode('#,##0');      
            $sheet->getStyle('K'.$currentRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
          }
        }
      
      }
    }

    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    $writer->save($outputFileName);

    return array('filename' => 'e-faktur.xlsx', 'filepath' => $outputFilePath);
  }  

  public function writeFileCSV($arrDataResult, $filename = "e_faktur.csv", $delimiter=",") 
  {
    $outputFileName = FCPATH.'assets/files-output/'.$filename;
    $outputFilePath = base_url().'assets/files-output/'.$filename;

    //create header 
    $arrHeader = array(
      array(
        'FK', 'KD_JENIS_TRANSAKSI', 'FG_PENGGANTI', 'NOMOR_FAKTUR', 'MASA_PAJAK', 'TAHUN_PAJAK', 'TANGGAL_FAKTUR',
        'NPWP', 'NAMA', 'ALAMAT_LENGKAP', 'JUMLAH_DPP', 'JUMLAH_PPN', 'JUMLAH_PPNBM', 'ID_KETERANGAN_TAMBAHAN', 
        'FG_UANG_MUKA', 'UANG_MUKA_DPP', 'UANG_MUKA_PPN', 'UANG_MUKA_PPNBM', 'REFERENSI', 'KODE_DOKUMEN_PENDUKUNG'
      ),
      array(
        'LT', 'NPWP', 'NAMA', 'JALAN', 'BLOK', 'NOMOR', 'RT', 'RW', 'KECAMATAN', 'KELURAHAN', 'KABUPATEN',
        'PROPINSI', 'KODE_POS', 'NOMOR_TELEPON', '', '', '', '', '', ''
      ),
      array(
        'OF', 'KODE_OBJEK', 'NAMA', 'HARGA_SATUAN', 'JUMLAH_BARANG', 'HARGA_TOTAL', 'DISKON', 'DPP', 'PPN',
        'TARIF_PPNBM', 'PPNBM', '', '', '', '', '', '', '', '', ''
      ),
    );

    if(!is_dir($outputFilePath)){
      mkdir($outputFilePath, 0755, true);
    }

    if(file_exists($outputFileName)){
      unlink($outputFileName);
    }

    header("Content-Disposition: attachment; filename=$filename"); 
    header("Content-Type: application/csv; ");
    header("Pragma: no-cache");
    header("Expires: 0");

    // file creation 
    $file = fopen($outputFileName, 'aw');
    if(!empty($arrHeader)){
      foreach($arrHeader AS $idx => $rowHead){
        fputs($file, implode($delimiter, $rowHead)."\n");
      }
    }

    //create arrdata for body csv
    if(!empty($arrDataResult))
    {
      foreach($arrDataResult AS $idx => $row)
      {
        //proses penentuan masa dan tahun pajak ambil dari bulan dan tahun dari invoice_date
        $tglFaktur  = "";
        $masaPajak  = "";
        $tahunPajak = "";
        if(!empty($row['invoice_date']))
        {
          $tglFaktur  = date("d/m/Y", strtotime($row['invoice_date']));
          $masaPajak  = date("m", strtotime($row['invoice_date']));
          $tahunPajak = date("Y", strtotime($row['invoice_date']));
        }

        $npwp = str_replace('.', '', $row['npwp']);
        $npwp = str_replace('-', '', $npwp);
        if ( $npwp == '' )
          $npwp = '000000000000000';

        $noFaktur = $row['nomor_faktur'];
        // $noFaktur = str_replace('.', '', $row['nomor_faktur']);
        // $noFaktur = str_replace('-', '', $noFaktur);
        // $noFaktur = substr($noFaktur,3);

        $alamat = str_replace(',', '', $row['alamat_lengkap']);
        $nama = str_replace(',', '', $row['nama']);


        $arrHeadBody = array(
          $row['code'], $row['jenis_transaksi'], $row['fp_pengganti'], $noFaktur, intval($masaPajak), intval($tahunPajak), 
          $tglFaktur, $npwp, strtoupper($nama), strtoupper($alamat), round($row['jumlah_dpp']),
          round($row['jumlah_ppn']), round($row['jumlah_ppnbm']), $row['id_keterangan_tambahan'], round($row['fg_uang_muka']),
          round($row['uang_muka_dpp']), round($row['uang_muka_ppn']), round($row['uang_muka_ppnbm']), $row['referensi'], ""
        );
        fputs($file, implode($delimiter, $arrHeadBody)."\n");
        
        if(!empty($row['detail']))
        {
          $arrDetail = $row['detail'];

          foreach($arrDetail AS $rowDetail)
          {
            $arrBody = array(
              $rowDetail['code'], $rowDetail['kode_objek'], $rowDetail['nama'], round(floatval($rowDetail['harga_satuan'])), $rowDetail['jumlah_barang'],
              $rowDetail['harga_total'], round(floatval($rowDetail['diskon'])), round(floatval($rowDetail['harga_dpp'])), round(floatval($rowDetail['ppn'])), round(floatval($rowDetail['tarif_ppnbm'])),
              round(floatval($rowDetail['ppnbm'])), 0, "", "", "", "", "", "", "", ""
            );
            fputs($file, implode($delimiter, $arrBody)."\n");
          }
        }
      
      }
    }

    fclose($file);
    return array('filename' => $filename, 'filepath' => $outputFilePath);
  }  
}
