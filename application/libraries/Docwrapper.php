<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

require FCPATH.'vendor'.DIRECTORY_SEPARATOR.'autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

Dompdf\Autoloader::register();
use Dompdf\Dompdf;

class Docwrapper {
	
	private $CI;
	private $header;
	private $body;
	private $data;
	private $option;
	private $style;

	private $doc_ex;

	public function __construct(){
		$this->CI =& get_instance();
	}

	public function make_header(){

	}

	public function set($column, $data = array(), $option = array()){

		foreach($column as $c){
			
			$this->header[]	= $c['label'];
			$this->body[]	= $c['data'];

		}

		$this->data 	= $data;
		$this->option	= $option;

		$this->processing_setup();

		return $this;
	}
	
	public function createExcel(){

		$excel = $this->processing_sheet();

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}
		
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
				
		$writer = IOFactory::createWriter($excel, 'Xls');
		$writer->save('php://output');
	}

	public function createHTML(){
		
		$excel = $this->processing_sheet();

		$excel->getActiveSheet()->getPageMargins()->setTop(0);
		$excel->getActiveSheet()->getPageMargins()->setRight(0);
		$excel->getActiveSheet()->getPageMargins()->setLeft(0);
		$excel->getActiveSheet()->getPageMargins()->setBottom(0);

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}

		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($excel);
		$html = $writer->generateHTMLHeader();
		$html .= $this->getStyleHTML();
		$html .= $writer->generateSheetData();
		$html .= $writer->generateHTMLFooter();

		echo $html;
		// $writer->save('php://output');
	}

	public function createPDF(){
		
		/*
		$excel = $this->processing_sheet();

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}

		$class = \PhpOffice\PhpSpreadsheet\Writer\Pdf\Dompdf::class;
		IOFactory::registerWriter( 'Pdf', $class );

		header( 'Content-Type: application/pdf' );
		header( 'Content-Disposition: attachment;filename="'.$filename.'.pdf"' );
		header( 'Cache-Control: max-age=0' );
		header( 'Cache-Control: max-age=1' );
		
		header( 'Expires: Mon, 26 Jul 1997 05:00:00 GMT' );
		header( 'Last-Modified: ' . gmdate( 'D, d M Y H:i:s' ) . ' GMT' );
		header( 'Cache-Control: cache, must-revalidate' );
		header( 'Pragma: public' );
		
		$writer = IOFactory::createWriter( $excel, 'Pdf' );
		$writer->save('php://output');
		*/



		$excel = $this->processing_sheet();

		$excel->getActiveSheet()->getPageMargins()->setTop(0);
		$excel->getActiveSheet()->getPageMargins()->setRight(0);
		$excel->getActiveSheet()->getPageMargins()->setLeft(0);
		$excel->getActiveSheet()->getPageMargins()->setBottom(0);

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}

		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($excel);
		$html = $writer->generateHTMLHeader();
		$html .= $this->getStyleHTML();
		$html .= $writer->generateSheetData();
		$html .= $writer->generateHTMLFooter();

		$dompdf = new Dompdf();
		$dompdf->loadHtml($html);
		$dompdf->setPaper('A4', 'landscape');
		$dompdf->render();
		$dompdf->stream($filename);
	}

	public function getStyleHTML(){
		
		$style['header_title'] = 'line-height: 1.42; font-weight: 700; font-size: 14pt; text-align: center;';
		$style['header'] = 'font-size: 108%; text-align: center;';
		$style['header_'] = 'border-bottom: 4px double black;';
		$style['header_doc'] = 'padding: 20px 30px 30px 30px; line-height: 1.42; text-align: center; font-weight: 700; font-size: 1.17em;';

		$style['table_header'] = 'padding: 6px 8px; border: 1px solid black !important;';
		$style['table_body'] = 'padding: 6px 8px; border: 1px solid black !important;';
		$style['summary'] = 'padding: 6px 8px; border: 1px solid black !important;';

		$return = '<link type="text/css" rel="stylesheet" href="'.base_url('css/report-print.css').'" />';
		$return .= '<style>';
		
		$return .= '@media print {
			@page {
				size: 297mm 210mm;
				margin: 2cm 2cm 1.5cm 1.5cm;
			}
		}';

		$return .= 'body{ margin: 8px !important; } table{ width: 100%; } table td{ border: initial; }';

		foreach($this->style as $k => $s){
			if(isset($style[$k])){
				$return .= implode(',', $s).'{'.$style[$k].'}';
			}
		}
		$return .= '</style>';

		return $return;
	}

	public function setStyleClass($type, $rownumber){
		return $this->style[$type][] = '.row'.($rownumber-1).' td';
	}

	public function processing_setup(){

		$column			= $this->header;
		$data_body		= $this->body;
		$alpha			= array();
		$alpha_first	= 'A';
		$alpha_last		= '';
		$column_pos		= array();

		for($i=1; $i<=count($column); $i++){
			$alpha[$i] = Coordinate::stringFromColumnIndex($i);
			$alpha_last = $alpha[$i];

			// $column_pos[$alpha[$i]] = $this->header[$i-1]['data'];
			$column_pos[$data_body[$i-1]] = $alpha[$i];
		}

		$this->doc_ex['index_alpha']	= $alpha;
		$this->doc_ex['first_alpha']	= $alpha_first;
		$this->doc_ex['last_alpha']		= $alpha_last;
		$this->doc_ex['index_total']	= count($column);
		$this->doc_ex['column_pos']		= $column_pos;

		$data			= $this->data;
		$formatted_data	= array();
		if(!empty($data)){
			foreach($data as $d){
				$result = array_intersect_key($d, array_flip($this->body));
				$formatted_data[] = array_replace(array_flip($this->body), $result);
			}
		}

		$this->data = $formatted_data;
	}

	public function processing_sheet(){

		$prop			= $this->doc_ex;
		$row_counter	= 1;
		$data			= $this->data;
		$merged			= (isset($this->option['merged']) && is_array($this->option['merged'])) ? $this->option['merged'] : false;

		$excel	= new Spreadsheet();
		$sheet	= $excel->setActiveSheetIndex(0);

		$this->get_header_excel($sheet, $prop, $row_counter);
		$this->get_filter_info_excel($sheet, $prop, $row_counter);
		
		$row_counter++;

		// header table
		
		$cell_body = $prop['first_alpha'].$row_counter;
		foreach($prop['index_alpha'] as $k => $a){
			$sheet->setCellValue($a.$row_counter, $this->header[$k-1]);
			$this->setStyleClass('table_header', $row_counter);
		}
		
		$row_counter++;
		$mark_group = '';
		$cell_group = array();
		
		// body table
		foreach($data as $d){
			
			$a = $prop['first_alpha'];
			
			foreach($d as $dk => $dv){
				if(isset($d[$merged['by']]) && !in_array($dk, $merged['except'])){

					if($mark_group != $d[$merged['by']]){
						$mark_group = $d[$merged['by']];
						$cell_group[$a.'_'.$mark_group][] = $a.$row_counter;
					}else{
						$cell_group[$a.'_'.$mark_group][] = $a.$row_counter;
					}
				}

				$this->setStyleClass('table_body', $row_counter);
				$sheet->setCellValue($a.$row_counter, $dv);
				$a++;
			}

			
			
			$row_counter++;
		}

		if(!empty($cell_group)){
			foreach($cell_group as $g){

				$g_first = reset($g);
				$g_last = end($g);

				if($g_first && $g_last){
					$sheet->mergeCells($g_first.':'.$g_last);
					$sheet->getStyle($g_first.':'.$g_last)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
				}
			}
		}

		$this->get_summary_excel($sheet, $prop, $row_counter);


		$cell_body .= ':'.$prop['last_alpha'].($row_counter-1);

		$styleArray = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ),
            ),
        );

        $sheet->getStyle($cell_body)->applyFromArray($styleArray);

		foreach($prop['index_alpha'] as $a){
            $sheet->getColumnDimension($a)->setAutoSize(true);
		}
		
		return $excel;
	}

	public function get_header_excel(&$sheet, $prop, &$row_counter){

		$blogo = get_parameter( 'report_business_logo', 'report_header_param' );
        $bname = get_parameter( 'report_business_name', 'report_header_param' );
        $bdet  = get_parameter( 'report_business_detail', 'report_header_param' );
		$ebdet = explode('<br>', $bdet['key_value_1']);

		$cell_centered = $prop['first_alpha'].$row_counter;
		
		$this->setStyleClass('header_title', $row_counter);
		$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
		$sheet->setCellValue($prop['first_alpha'].$row_counter++, $bname['key_value_1']);
		
		
		$this->setStyleClass('header', $row_counter);
		$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
		$sheet->setCellValue($prop['first_alpha'].$row_counter++, $ebdet[0]);
		
		$this->setStyleClass('header', $row_counter);
		$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
		$sheet->setCellValue($prop['first_alpha'].$row_counter++, $ebdet[1]);
		
		$this->setStyleClass('header', $row_counter);
		$this->setStyleClass('header_', $row_counter);
		$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
		$sheet->getStyle($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE);
		$sheet->setCellValue($prop['first_alpha'].$row_counter++, $ebdet[2]);
		
		
		if(isset($this->option['title'])){
			$row_counter++;
			$this->setStyleClass('header_doc', $row_counter);
			$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
			$sheet->setCellValue($prop['first_alpha'].$row_counter++, $this->option['title']);
		}

		$cell_centered .= ':'.$prop['last_alpha'].$row_counter;

		$sheet->getStyle($cell_centered)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
	}

	public function get_filter_info_excel(&$sheet, $prop, &$row_counter){

		if(!isset($this->option['filter']) || !is_array($this->option['filter'])){
			return;
		}

		$row_counter++;
		
		foreach($this->option['filter'] as $k => $f){
			$this->setStyleClass('filter', $row_counter);
			$sheet->setCellValue($prop['first_alpha'].$row_counter++, $f['label'].' : '.$f['value']);
		}

		$row_counter++;

	}

	public function get_summary_excel(&$sheet, $prop, &$row_counter)
	{
		if(!isset($this->option['summary']) || !is_array($this->option['summary'])){
			return;
		}

		foreach($this->option['summary'] as $f){
			$this->setStyleClass('summary', $row_counter);
			$sheet->setCellValue($prop['first_alpha'].$row_counter, $f['label']);

			if(!is_array($f['column'])){
				if(isset($prop['column_pos'][$f['column']])){
					$sheet->setCellValue($prop['column_pos'][$f['column']].$row_counter, $f['value']);
				}
			}else{
				foreach($f['column'] as $ck => $cv){
					$sheet->setCellValue($prop['column_pos'][$ck].$row_counter, $cv);
				}
			}
			
			$row_counter++;
		}
	}
}