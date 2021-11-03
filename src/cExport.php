<?php
require_once('config.php');
require_once('Class.MySQL.Lite.php');
require_once('class.Wholesaler.php');
require_once('PHPExcel.php');

//Clear Cache
header("Cache-Control: no-cache, must-revalidate"); // HTTP/1.1
header("Expires: Sat, 26 Jul 1997 05:00:00 GMT"); // Date in the past

/**
* Import Class
*
* Contains all functionality related to importing orders
*
* @author Elemental
* @version 1.0
* @copyright The Wholesaler
*/
class cExport
{
    public $db;
    /**
     * @ignore
     */
    public function __construct()
    {
        global $site;
        //initialize DB logic
        $this->db  = new MySQL_Lite_ABS($site['db']);
		$this->app = new Wholesaler();
		$this->xls = new PHPExcel();
    }

	public function export_pick_list($order_num)
	{
		$l = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

		//Get customer details from order number
		$sql = "SELECT
		        tbl_registrations.ta_company_name,
		        tbl_registrations.phone,
				IF (orders.custom_address IS NOT NULL AND orders.custom_address != '',orders.custom_address,tbl_registrations.del_add) AS del_add,
				IF (orders.custom_address2 IS NOT NULL AND orders.custom_address2 != '',orders.custom_address2,tbl_registrations.del_add2) AS del_add2,
				IF (orders.custom_address_city IS NOT NULL AND orders.custom_address_city != '',orders.custom_address_city,tbl_registrations.del_add_city) AS del_add_city,
				IF (orders.custom_address_zip IS NOT NULL AND orders.custom_address_zip != '',orders.custom_address_zip,tbl_registrations.del_add_zip) AS del_add_zip,
				IF (orders.custom_address_phone IS NOT NULL AND orders.custom_address_phone != '',orders.custom_address_phone,tbl_registrations.del_add_phone) AS del_add_phone,
		        orders.outcome_comments,
		        orders.freight_type,
		        orders.courier,
				orders.promo_id
		        FROM  orders
		        LEFT JOIN tbl_registrations ON tbl_registrations.id = orders.user_reg_id
		        WHERE orders.ord_id = '".$order_num."'";
		$customer_details = $this->db->RetrieveCommandExec($sql);

		//Get Full Name for Freight Type
		/*
		if ($customer_details[0]['freight_type'] == "collect")
		{
			$freight_type = "Collect";
		}
		elseif($customer_details[0]['courier'] != "" && $customer_details[0]['courier'] != 0)
		{
			$freight_type = $customer_details[0]['courier'];
		}
		elseif ($customer_details[0]['freight_type'] == "customer_courier")
		{
			$freight_type = "Customer's Courier";
		}
		elseif ($customer_details[0]['freight_type'] == "our_choice")
		{
			$freight_type = !empty($customer_details[0]['courier']) ? ucfirst(str_replace("_", " ", $customer_details[0]['courier'])) : "Aramex";
		}
		elseif ($customer_details[0]['freight_type'] == "our_choice_aramex")
		{
			$freight_type = !empty($customer_details[0]['courier']) ? ucfirst(str_replace("_", " ", $customer_details[0]['courier'])) : "Aramex";
		}
		elseif ($customer_details[0]['freight_type'] == "our_choice_courier_guy")
		{
			$freight_type = !empty($customer_details[0]['courier']) ? ucfirst(str_replace("_", " ", $customer_details[0]['courier'])) : "The Courier Guy";
		}
		else
		{
			$freight_type = "Deliver";
		}*/

		$freight_type = "Deliver";
        if($customer_details[0]['freight_type'] == "aramex" || $customer_details[0]['freight_type'] == "our_choice") {
			$freight_type = "Aramex";
        } else if($customer_details[0]['freight_type'] == "courier_guy") {
			$freight_type = "The Courier Guy";
        } else if($customer_details[0]['freight_type'] == "collect") {
			$freight_type = "Collect";
        } else if(!empty($customer_details[0]['courier'])) {
			$freight_type = "Collect - ".ucfirst(str_replace("_", " ", $customer_details[0]['courier']));
        }

		//Get order details
		$con =  mysqli_connect("dedi109.cpt3.host-h.net", "thewhg_12", "6?^%-pHpDUf_%PAk");
		if(mysqli_select_db($con, "thewhg_db12"))
		{
			$sql = "SELECT
					'' AS checkfield,
					'' AS comments,
					products.prod_name,
					products.prod_description,
					CONCAT(products.barcode,' ') AS barcode,
					order_details.price,
					order_details.qty,
					order_details.lots as 'prod_lots',
					order_details.ord_num,
					order_details.prod_id
					FROM
					order_details
					INNER JOIN products ON products.id = order_details.prod_id
					WHERE order_details.ord_num = '".mysqli_real_escape_string($con, $order_num)."'
					ORDER BY products.prod_name ASC, products.prod_description ASC";
			mysqli_close($con);
		}
		$result = $this->db->RetrieveCommandExec($sql);

		$fields = array(
			'checkfield'        => 'Check',
			'qty'               => 'Qty',
			'barcode'           => 'Barcode',
			'prod_name'         => 'Product Name',
			'prod_description'  => 'Description',
			'prod_lots'         => 'Lots',
			'comments'          => 'Comments'
		);

		//$this->standard_export($result,$fields,'pick_list_'.$order_num);

		$this->xls = PHPExcel_IOFactory::createReader('Excel5');
		$this->xls = $this->xls->load('../xls_templates/picklist_template.xls');

		//Excel Properties
		$this->xls->getProperties()->setCreator("The Wholesaler")
			->setLastModifiedBy("The Wholesaler")
			->setTitle("The Wholesaler - ".date("Y/m/d"))
			->setSubject("The Wholesaler - ".date("Y/m/d"))
			->setDescription("The Wholesaler");

		$sheet = $this->xls->setActiveSheetIndex(0);

		// set default values
		$address1 = "";
		$address2 = "";
		$city = "";
		$zip = "";
		$phone = "";

		// perform empty checks
		if(!empty($customer_details[0]['del_add'])){
			$address1 = $customer_details[0]['del_add'];
		}
		if(!empty($customer_details[0]['del_add2'])){
			$address2 = $customer_details[0]['del_add2'];
		}
		if(!empty($customer_details[0]['del_add_city'])){
			$city = $customer_details[0]['del_add_city'];
		}
		if(!empty($customer_details[0]['del_add_zip'])){
			$zip = $customer_details[0]['del_add_zip'];
		}
		if(!empty($customer_details[0]['del_add_phone'])){
			$phone = $customer_details[0]['del_add_phone'];
		}

		//Set customer information
		$sheet->setCellValue('D3', date('d M, Y'));
		$sheet->setCellValue('D4', $order_num.($customer_details[0]['promo_id'] > 0?'#'.str_pad($customer_details[0]['promo_id'], 4, "0", STR_PAD_LEFT):''));
		$sheet->setCellValue('D5', $customer_details[0]['ta_company_name']);
		$sheet->setCellValue('D6', $customer_details[0]['phone']);
		$sheet->setCellValue('D7', $address1);
		$sheet->setCellValue('D8', $address2);
		$sheet->setCellValue('D9', $city);
		$sheet->setCellValue('D10', $zip);
		$sheet->setCellValue('D11', $phone);
		$sheet->setCellValue('D12', $freight_type);

		//Set data
		$r = 15;
		foreach ($result as $row)
		{
			$i = 0;
			foreach ($fields as $field => $heading)
			{
				$sheet->setCellValue($l[$i].$r, trim($row[$field]));
				$i++;
			}

			//Every second row should be light grey
			if ($r % 2 == 0)
			{
				$sheet->getStyle("A{$r}:G{$r}")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
				$sheet->getStyle("A{$r}:G{$r}")->getFill()->getStartColor()->setARGB('FFEEEEEE');
			}
			else
			{
				$sheet->getStyle("A{$r}:G{$r}")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
				$sheet->getStyle("A{$r}:G{$r}")->getFill()->getStartColor()->setARGB('FFFFFFFF');
			}

			$r++;
		}

		//Autosize cells to fit content
		$i = 0;
		foreach ($fields as $heading)
		{
			$sheet->getColumnDimension($l[$i])->setAutoSize(true);
			$i++;
		}

		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="pick_list_'.$order_num.'.xls"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel5');
		$objWriter->save('php://output');
		exit();
	}

	public function standard_export($data, $fields, $filename, $template='')
	{
        // ini_set('memory_limit','256M');
		$l = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

		!empty($template) || $template = '../xls_templates/basic_template.xls';

		$this->xls = PHPExcel_IOFactory::createReader('Excel5');
		$this->xls = $this->xls->load($template);


		//Excel Properties
		$this->xls->getProperties()->setCreator("The Wholesaler")
			->setLastModifiedBy("The Wholesaler")
			->setTitle("The Wholesaler - ".date("Y/m/d"))
			->setSubject("The Wholesaler - ".date("Y/m/d"))
			->setDescription("The Wholesaler");

		$sheet = $this->xls->setActiveSheetIndex(0);

		//Set headings
		$i = 0;
		foreach ($fields as $heading)
		{
			$sheet->setCellValue($l[$i].'1', $heading);
			$i++;
		}
        // echo '<pre>';
        // print_r($data);die;
		//Set data
		$r = 2;
		foreach ($data as $row)
		{
			$i = 0;
			foreach ($fields as $field => $heading)
			{
				$sheet->setCellValue($l[$i].$r, trim($row[$field]));
				$i++;
			}

			//Every second row should be light grey
			if ($r % 2 == 0)
			{
				$sheet->getStyle("A{$r}:Z{$r}")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
				$sheet->getStyle("A{$r}:Z{$r}")->getFill()->getStartColor()->setARGB('FFEEEEEE');
			}

			$r++;
		}

		//Autosize cells to fit content
		$i = 0;
		foreach ($fields as $heading)
		{
			$sheet->getColumnDimension($l[$i])->setAutoSize(true);
			$i++;
		}

		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel5');
		$objWriter->save('php://output');
		exit();
	}

//End class
}
?>