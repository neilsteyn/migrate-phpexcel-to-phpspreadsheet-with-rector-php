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
class cImport
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

	/**
	 * Export Order List
	 *
	 * Creates an excel spreadsheet of all the active products on the system.
	 */
    public function exportOrderList()
	{
		//Excel Properties
		$this->xls->getProperties()->setCreator("The Wholesaler")
							 ->setLastModifiedBy("The Wholesaler")
							 ->setTitle("The Wholesaler - Products - ".date("Y/m/d"))
							 ->setSubject("The Wholesaler - Products - ".date("Y/m/d"))
							 ->setDescription("Product list for use with the 'Import Order' feature");

		//Get list of products
		$sql = "SELECT prod_name, prod_description, prod_add_date,
				(SELECT type FROM prod_types
					LEFT JOIN pivot_product_type ON pivot_product_type.type_id = prod_types.id
					WHERE pivot_product_type.product_id = products.id LIMIT 1) AS prod_type,
				prod_lots, prod_price, prod_retailPrice, prod_image,
				CONCAT(barcode,' ') AS barcode
				FROM products
				WHERE comingsoon = '0' AND inactive = 0 ORDER BY TRIM(prod_name) ASC";
		$result = $this->db->RetrieveCommandExec($sql);

		//Set headings
		$this->xls->setActiveSheetIndex(0)
				->setCellValue('A1', "Quantity")
				->setCellValue('B1', "Barcode")
				->setCellValue('C1', "Code")
				->setCellValue('D1', "Description")
				->setCellValue('E1', "Image")
				->setCellValue('F1', "Type")
				->setCellValue('G1', "Lots")
				->setCellValue('H1', "Price")
				->setCellValue('I1', "RRP")
				->setCellValue('J1', "Date Loaded");

		//Add bottom border to header row
		$this->xls->getActiveSheet()->getStyle("A1:J1")->applyFromArray(
			array(
				'borders' => array(
					'bottom'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$i=2;
		foreach ($result as $row)
		{
			$image = str_replace('../', ADDRESS, $row['prod_image']);
			$image = str_replace('images/prods/', 'images/prods/big/', $image);
			//Set cell values
			$this->xls->setActiveSheetIndex(0)
				->setCellValue('B'.$i, trim($row['barcode']))
				->setCellValue('C'.$i, trim($row['prod_name']))
				->setCellValue('D'.$i, trim($row['prod_description']))
				->setCellValue('E'.$i, trim($image))
				->setCellValue('F'.$i, trim($row['prod_type']))
				->setCellValue('G'.$i, trim($row['prod_lots']))
				->setCellValue('H'.$i, trim($row['prod_price']))
				->setCellValue('I'.$i, trim($row['prod_retailPrice']))
				->setCellValue('J'.$i, trim($row['prod_add_date']));

			$this->xls->setActiveSheetIndex(0)->getCell('D'.$i)
				->getHyperlink($image)
				->setUrl($image);

			//Every second row should be light grey
			if ($i % 2 == 0)
			{
				$this->xls->getActiveSheet()->getStyle("A{$i}:J{$i}")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
				$this->xls->getActiveSheet()->getStyle("A{$i}:J{$i}")->getFill()->getStartColor()->setARGB('FFEEEEEE');
			}

			$i++;
		}

		$i--;
		//add a border on the left side
		$this->xls->getActiveSheet()->getStyle("B1:B{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("C1:C{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("D1:D{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("E1:E{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("F1:F{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("G1:G{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("H1:H{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("I1:I{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		$this->xls->getActiveSheet()->getStyle("J1:J{$i}")->applyFromArray(
			array(
				'borders' => array(
					'left'     => array(
						'style' => PHPExcel_Style_Border::BORDER_THIN
					)
				)
			)
		);

		//Format price column
		$this->xls->getActiveSheet()->getStyle("H2:H{$i}")->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
		$this->xls->getActiveSheet()->getStyle("I2:I{$i}")->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);

		$this->xls->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
		$this->xls->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);

		$this->xls->getActiveSheet()->getStyle('A1:J1')->getFont()->setBold(true);

		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="products_list.xls"');
		header('Cache-Control: max-age=0');

//		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//		header('Content-Disposition: attachment;filename="products_list.xlsx"');
//		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel5');
//		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel2007');
		$objWriter->save('php://output');
		exit();
		echo "Done";
	}

	/**
	 * Validate Import File
	 *
	 * Validates the imported file
	 *
	 * @return string
	 */
	public function validate_import_file()
	{
		$error = "";

		if ($_FILES['file']['error'] == 4)
		{
			$error = "No file selected!";
		}
		elseif ($_FILES['file']['type'] != "application/vnd.ms-excel" && $_FILES['file']['type'] != "text/csv" && $_FILES['file']['type'] != "application/octet-stream")
		{
			$error = "Invalid file type!";
		}
		else
		{
			if (is_uploaded_file($_FILES["file"]["tmp_name"]))
			{
				//Set file path
				$filepath	= "../admin/temp/";
				$file		= $_SESSION['userid']."_".$_FILES["file"]["name"];

				//Precaution: Delete the file if it already exists in the folder
				if (file_exists($filepath.$file)) unlink($filepath.$file);

				//Move file to temp directory
				if (move_uploaded_file($_FILES["file"]["tmp_name"], $filepath.$file))
				{
					//Import file
					$ext = substr($_FILES["file"]["name"], -3);
					if ($ext == "xls")
					{
						$error = $this->import_xls_order($filepath.$file);
					}
					else
					{
						$error = $this->import_csv_order($filepath.$file);
					}

					if (empty($error))
					{
						//Redirect to confirm order page
						header("Location: ".ADDRESS."Confirm_import");
					}
				}
			}
			else
			{
				$error = "Error uploading file";
			}
		}

		return $error;
	}

    /**
	 * Import xls Order
	 *
	 * Import the products from an xls file
	 * Values will be stored in the import session
	 *
	 * @param string $file path/file.ext
	 */
	public function import_xls_order($file)
	{
		//unset old import session if it still exists
		if (isset($_SESSION['import'])) unset($_SESSION['import']);

		require_once 'PHPExcel/IOFactory.php';
		$objPHPExcel = PHPExcel_IOFactory::load($file);

		$objWorksheet = $objPHPExcel->setActiveSheetIndex('0');

		//Run through each row
		$highestRow = $objWorksheet->getHighestRow();
		for ($row = 1; $row <= $highestRow; $row++)
		{
			$qty = $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();

			//if qty isn't empty/null/0 and is a valid number
			if (!empty($qty) && $qty != null && (int)$qty != 0 && is_numeric($qty))
			{
				$code = $objWorksheet->getCellByColumnAndRow(2, $row)->getValue();
			
				$code	= str_replace(" ", "", $code);
				$code	= str_replace("-", "", $code);

				//Do not add the product if it's already been added
				if (isset($_SESSION['import']))
				{
					if(!array_key_exists($code, $_SESSION['import']))
					{
						$_SESSION['import'][$code] = (int)$qty;
					}
				}
				else
				{
					$_SESSION['import'][$code] = (int)$qty;
				}
			}
		}

		if (!isset($_SESSION['import'])) return "There were no products to import (xls)";

//		print_r($_SESSION['import']);
	}

	/**
	 * Import csv Order
	 *
	 * Import the products from an csv file
	 * Values will be stored in the import session
	 *
	 * @param string $file path/file.ext
	 */
	public function import_csv_order($file)
	{
		//unset old import session if it still exists
		if (isset($_SESSION['import'])) unset($_SESSION['import']);

		$handle = fopen($file, "r");
		$i=0;

		$separator = $this->app->determine_separator($file);

		while (($csv = fgetcsv($handle, 50000, $separator)) !== FALSE)
		{
			$qty	= &$csv[0];

			//if qty is a valid number
			if (is_numeric($qty))
			{

				
				$code	= &$csv[1];
				

				if(preg_match("/[a-z]/i", $code)){
					$code	= str_replace(" ", "", $code);
					$code	= str_replace("-", "", $code);
				}else{
					$code   = trim(strip_tags((int)$code));
				}
			
				//Do not add the product if it's already been added
				if (isset($_SESSION['import']))
				{
					if(!array_key_exists($code, $_SESSION['import']))
					{
						$_SESSION['import'][$code] = (int)$qty;
					
					}
				}
				else
				{
					$_SESSION['import'][$code] = (int)$qty;
					
				}

			}
		}

		if (!isset($_SESSION['import'])) return "There were no products to import (csv)";
	}

	/**
	 * Process Import
	 *
	 * Processes the imported products,
	 * checkin if there is enough stock and if it's a valid product on the system
	 *
	 * @return array
	 */
	function process_import()
	{
		$i=0;
		$j=0;
		$total=0;
		$products = array();
		$loaded = array();

		if (isset($_SESSION['import']))
		{
			//Loop through each imported item
			foreach ($_SESSION['import'] as $code => $qty)
			{

				if(is_numeric($code)){
					$sql = "SELECT * FROM products
					WHERE comingsoon = '0'
					AND barcode = '{$code}'
					AND inactive = 0";
				}else{
					//Only add products which are currently available
					$sql = "SELECT * FROM products
					WHERE comingsoon = '0'
					AND REPLACE(REPLACE(prod_name, '-', ''), ' ', '') = '{$code}'
					AND inactive = 0";
				}
				
				$result = $this->db->RetrieveCommandExec($sql);

				if (!empty($result))
				{
					$products[$i]['qty']			= $qty;
					$products[$i]['id']				= $result[0]['id'];
					$products[$i]['image']			= $result[0]['prod_image'];
					$products[$i]['description']	= $result[0]['prod_name']." - ".$result[0]['prod_description'];
					$products[$i]['lots']			= $result[0]['prod_lots'];
					$products[$i]['price']			= $result[0]['prod_price'];

					//Flag if item is out of stock
					$products[$i]['instock']		= $result[0]['prod_stock'] > $qty ? true : false;

					//Accumulate Total
					$total += ($qty * $result[0]['prod_lots'] * $result[0]['prod_price']);

					//Add item as valid product to be imported to cart once approved
					$_SESSION['valid_product'][$i]['id']		= $result[0]['id'];
					$_SESSION['valid_product'][$i]['qty']		= $qty;
					$_SESSION['valid_product'][$i]['price']		= $result[0]['prod_price'];
					$_SESSION['valid_product'][$i]['lots']		= $result[0]['prod_lots'];

					$i++;
				}
				else
				{
					//These items doesn't exist or has been removed
					$invalid[$j]['qty']				= $qty;
					$invalid[$j]['description']		= $code;
					$j++;
				}
			}
		}
		//Create import array
		$import = array('products'=>$products,
						'total'=>$total,
						'invalid'=>$invalid);
		return $import;
	}

	function import_to_cart()
	{
		//Manual users should have their own order number
		$manual = $_SESSION['user_type'] == "MANUAL" ? "1" : "0";

		$ord_num = '';

		$cmd = "SELECT
				ord_id
				FROM
				orders
				WHERE
				user_reg_id = " . $this->app->CheckInjection($_SESSION['userid']) . "
				AND order_status = 'PENDING'
				AND manual = '{$manual}'
				";

		$result = $this->db->RetrieveCommandExec($cmd);

		$i=0;
		$order['id'] = array();
		if(empty($result))
		{
			//if there is no pending order then genereate a new order number
			if(!$ord_num = $this->app->GenerateOrderNumber())
			{
				return false;
			}
			else
			{
				//create the entry in the database
				$cmd = "INSERT INTO orders(ord_id,user_reg_id,order_status,order_date,manual)
						VALUES
						(
							'" . $ord_num . "',
							'" . $this->app->CheckInjection($_SESSION['userid']) . "',
							'PENDING',
							NOW(),
							'{$manual}'
						)";
				if(!($this->db->ActionCommandExec($cmd))) $this->app->reportDBError("import_to_cart", $cmd);

			}
		}
		else
		{
			$ord_num = $result[0]['ord_id'];

			//If the order session is set add all items the order array
			if (!empty($_SESSION['order']))
			{
				//var_dump($_SESSION['order']); exit();
				foreach($_SESSION['order']['id'] as $prodid)
				{
					$order['id'][$i]		= $prodid;
					$order['quantity'][$i]	= $_SESSION['order']['quantity'][$i];
					$order['price'][$i]		= $_SESSION['order']['price'][$i];
					$i++;
				}
			}
		}

		//Start sql query for bulk insert
		$sql = "";
		foreach ($_SESSION['valid_product'] as $row)
		{
			//If the product already exists on the order, then just increase the qty
			$key = array_search($row['id'], $order['id']);
			if ($key !== false)
			{
				$order['quantity'][$key] += $row['qty'];

				//update the order details table
				$cmd = "UPDATE order_details SET qty = ".$order['quantity'][$key]." WHERE ord_num = '".$ord_num."' AND prod_id = ".$order['id'][$key];
				if(!($this->db->ActionCommandExec($cmd))) $this->app->reportDBError("import_to_cart", $cmd);
			}
			else
			{
				//Add values for bulk insert
				$sql .= "('".$ord_num."', ".$row['id'].", ".$row['price'].", ".$row['qty'].", ".$row['lots']."),";

				//Build array in the format of the order session array
				$order['id'][$i]		= $row['id'];
				$order['quantity'][$i]	= $row['qty'];
				$order['price'][$i]		= $row['price'];
				$i++;
			}
		}

		// if $sql is empty then there are no new products to add to the cart
		if (!empty($sql))
		{
			$sql = rtrim($sql, ",");
			$sql = "INSERT INTO order_details (ord_num, prod_id, price, qty, lots) VALUES" . $sql;

			if(!($this->db->ActionCommandExec($sql)))
			{
				$this->app->reportDBError("import_to_cart", $sql);
			}
		}
		//Success
		$_SESSION['order'] = $order;
//		exit('end');
	}

	/**
	 * Import xls Products
	 *
	 * Import system used for importing the products for the new
	 * category layout created in Phase 25.
	 * Might get used in the future.
	 *
	 * @param string $file path/file.ext
	 */
	public function import_xls_products($file)
	{
		require_once 'PHPExcel/IOFactory.php';
		$objPHPExcel = PHPExcel_IOFactory::load($file);

		$objWorksheet = $objPHPExcel->setActiveSheetIndex('0');

		//Run through each row
		$highestRow = $objWorksheet->getHighestRow();
		for ($row = 2; $row <= $highestRow; $row++)
		{
			//Set columns
			$id				= $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
			$prod_name		= $objWorksheet->getCellByColumnAndRow(1, $row)->getValue();
			$prod_desc		= $objWorksheet->getCellByColumnAndRow(2, $row)->getValue();
			$inactive		= $objWorksheet->getCellByColumnAndRow(3, $row)->getValue();
			$brand_id		= $objWorksheet->getCellByColumnAndRow(4, $row)->getValue();
			$col_id			= $objWorksheet->getCellByColumnAndRow(5, $row)->getValue();
			$type_id		= $objWorksheet->getCellByColumnAndRow(6, $row)->getValue();

			$sql = 'UPDATE products
					SET prod_name = "'.$prod_name.'",
					prod_description = "'.$prod_desc.'",
					inactive = '.($inactive==""?0:$inactive).',
					brand_id = '.($brand_id==""?0:$brand_id).',
					collection_id = '.($col_id==''?0:$col_id).'
					WHERE id = '.$id;

			if ($this->db->ActionCommandExec($sql))
			{
				//multiple types are seperated by commas
				$types = explode(",", $type_id);

				foreach ($types as $t)
				{
					if (!empty($t))
					{
						//Delete old entries first incase we have to do it again
						$sql = "DELETE FROM pivot_product_type WHERE product_id = {$id} AND type_id = {$t}";
						if ($this->db->ActionCommandExec($sql))
						{
							$sql = "INSERT INTO pivot_product_type (product_id, type_id)
									VALUES ({$id}, {$t})";
							if (!$this->db->ActionCommandExec($sql))
							{
								die("Failed to insert into pivot at ID: {$id} <br/>SQL: {$sql}");
							}

						}
						else
						{
							die("Failed to delete from pivot at ID: {$id} <br/>SQL: {$sql}");
						}
					}
				}
			}
			else
			{
				die("Failed to update product at ID: {$id} <br/>SQL: {$sql}");
			}
		}

		echo "Import Successful!";
	}

	public function import_xls_core_products($file)
	{
		require_once 'PHPExcel/IOFactory.php';
		$objPHPExcel = PHPExcel_IOFactory::load($file);

		$objWorksheet = $objPHPExcel->setActiveSheetIndex('0');

		//Run through each row
		$highestRow = $objWorksheet->getHighestRow();
		for ($row = 2; $row <= $highestRow; $row++)
		{
			//Set columns
			$id				= $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
			$prod_name		= $objWorksheet->getCellByColumnAndRow(1, $row)->getValue();
			$prod_desc		= $objWorksheet->getCellByColumnAndRow(2, $row)->getValue();
			$inactive		= $objWorksheet->getCellByColumnAndRow(3, $row)->getValue();
			$core			= $objWorksheet->getCellByColumnAndRow(4, $row)->getValue();

			$sql = 'UPDATE products
					SET prod_name = "'.$prod_name.'",
					prod_description = "'.$prod_desc.'",
					inactive = '.($inactive==""?0:$inactive).',
					is_core = '.($core==""?0:$core).'
					WHERE id = '.$id;

			if (!$this->db->ActionCommandExec($sql))
			{
				die("Failed to update product at ID: {$id} <br/>SQL: {$sql}");
			}
		}

		echo "Import Successful!";
	}

	public function import_products_to_live($file)
	{
		require_once 'PHPExcel/IOFactory.php';
		$objPHPExcel = PHPExcel_IOFactory::load($file);

		$objWorksheet = $objPHPExcel->setActiveSheetIndex('0');

		//Run through each row
		$highestRow = $objWorksheet->getHighestRow();
		for ($row = 2; $row <= $highestRow; $row++)
		{
			//Set columns
			$id				= $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
			$prod_name		= $objWorksheet->getCellByColumnAndRow(1, $row)->getValue();
			$prod_desc		= $objWorksheet->getCellByColumnAndRow(2, $row)->getValue();
			$inactive		= $objWorksheet->getCellByColumnAndRow(3, $row)->getValue();
			$core			= $objWorksheet->getCellByColumnAndRow(4, $row)->getValue();
			$brand_id		= $objWorksheet->getCellByColumnAndRow(5, $row)->getValue();
			$col_id			= $objWorksheet->getCellByColumnAndRow(6, $row)->getValue();

			$sql = 'UPDATE products
					SET prod_name = "'.$prod_name.'",
					prod_description = "'.$prod_desc.'",
					inactive = '.($inactive==""?0:$inactive).',
					brand_id = '.($brand_id==""?0:$brand_id).',
					collection_id = '.($col_id==''?0:$col_id).',
					is_core = '.($core==""?0:$core).'
					WHERE id = '.$id;

			if (!$this->db->ActionCommandExec($sql))
			{
				die("Failed to update product at ID: {$id} <br/>SQL: {$sql}");
			}
		}

		echo "Import Successful!";
	}

	//
	public function update_brands($file)
	{
		require_once 'PHPExcel/IOFactory.php';
		$objPHPExcel = PHPExcel_IOFactory::load($file);

		$objWorksheet = $objPHPExcel->setActiveSheetIndex('0');

		//Run through each row
		$highestRow = $objWorksheet->getHighestRow();
		for ($row = 2; $row <= $highestRow; $row++)
		{
			//Set columns
			$id				= $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
			$brand_id		= $objWorksheet->getCellByColumnAndRow(1, $row)->getValue();

			$sql = 'UPDATE products SET
					brand_id = '.($brand_id==""?0:$brand_id).'
					WHERE id = '.$id;

			if (!$this->db->ActionCommandExec($sql))
			{
				die("Failed to update product at ID: {$id} <br/>SQL: {$sql} <br/>Row: {$row}");
			}
		}

		echo "New brands have successfully been assigned to all products";
	}

//End class
}
?>