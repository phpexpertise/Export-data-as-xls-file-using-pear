<?php
/**
 * Author : phpexpertise <info@phpexpertise.com>
 * Load config file and excel library file.
 */
class ExcelPear{
    protected $db;
    
    public function __construct($db_connect){        
        $this->db    =    $db_connect;
    }
    
    public function setTitle($name){
        $this->title = $name;
    }
    
    public function getTitle(){
        return $this->title;
    }
	
	public function getStatus($status_id){
        return ($status_id==1)?"Active":"Deactive";
    }
    /**
	 * Clear the spreadsheet tmp files.
	 * @params - none	 
	 */
    public function clearSpreadsheetCache() {
        $files = glob(BASEPATH.'/tmp/'.'*');        
        if ($files) {
            foreach ($files as $file) {
                if (file_exists($file)) {
                    @unlink($file);
                }
            }
        }
    }
    /**
	 * Get the users information only active users list
	 * @params - none	 
	 */
    public function get_users_results(){
        try{
            $stmt    =    $this->db->prepare("select * from users where status = :status");
            $stmt->execute(array(':status'=>1));
            $userRows    =    $stmt->fetchAll(PDO::FETCH_OBJ);
            if($stmt->rowCount()>0)
                return $userRows;
            else
                return false;
        }
        catch(PDOException $e){
            $e->getMessage();
        }
    }
    /**
	 * Export the users records into excel file using pear library
	 * @params - none	 
	 */
    public function export_users(){
        // LOAD THE PEAR LIBRARY
        $dir    =    BASEPATH.'\pear';        
        chdir($dir);
        require_once "Spreadsheet/Excel/Writer.php";        
        chdir(BASEPATH);
        
        $workbook = new Spreadsheet_Excel_Writer(); // Creating a workbook
        
        $workbook->setTempDir(BASEPATH.'\tmp');
        $workbook->setVersion(8);
                
        $box_format_array = array(
                            'Size' => 12,
                            'vAlign' => 'vequal_space',
                            'bold'=>1,
                            'Align'=>'center'
                            );
        
        $text_format_array = array(
                            'Size' => 12,
                            'vAlign' => 'vequal_space',
                            'Align'=>'center'
                            );

        $priceFormat    =& $workbook->addFormat(array('Size' => 10,'Align' => 'right','NumFormat' => '######0.00'));
        $boxFormat      =& $workbook->addFormat($box_format_array);            
        $boxFormat->setFontFamily('Calibri');
        $weightFormat   =& $workbook->addFormat(array('Size' => 10,'Align' => 'right','NumFormat' => '##0.00'));
        $textFormat     =& $workbook->addFormat($text_format_array);        
        $textFormat->setFontFamily('Calibri');
        
        $workbook->send($this->getTitle().'.xls'); // sending HTTP headers
        
        // Create worksheet
        $worksheet =& $workbook->addWorksheet(Ucfirst($this->getTitle()));
        $worksheet->setInputEncoding ( 'UTF-8' );
        $this->exportUsers( $worksheet,$this->db,$boxFormat, $textFormat );
        $worksheet->freezePanes(array(1, 0, 1, 1));
        
        // Close the workbook
        $workbook->close();
        
        // Clear Spreadsheet tmp files
        $this->clearSpreadsheetCache();
    }
    public function exportUsers(&$worksheet, &$database,&$boxFormat, &$textFormat){
        // Set column width
        $j = 0;
        $worksheet->setColumn($j,$j++,10);            
        $worksheet->setColumn($j,$j++,25);            
        $worksheet->setColumn($j,$j++,30);            
        $worksheet->setColumn($j,$j++,25);            
        $worksheet->setColumn($j,$j++,25);            
        $worksheet->setColumn($j,$j++,30);            
        // HEAER ROW
        $i = 0;
        $j = 0;
        $worksheet->writeString( $i, $j++, 'User ID', $boxFormat );        
        $worksheet->writeString( $i, $j++, 'User Name', $boxFormat );        
        $worksheet->writeString( $i, $j++, 'Email', $boxFormat );        
        $worksheet->writeString( $i, $j++, 'password', $boxFormat );        
        $worksheet->writeString( $i, $j++, 'Status', $boxFormat );                
        $worksheet->setRow( $i, 20, $boxFormat );
        $i += 1;
        $j = 0;
        
        $users_results    =    $this->get_users_results();
        if($users_results){
            $k=1;
            foreach ($users_results as $users_result) {
                $worksheet->setRow( $i, 20 );
                $worksheet->writeString($i, $j++, $k,$textFormat);
                $worksheet->writeString($i, $j++, html_entity_decode($users_result->username,ENT_QUOTES,'UTF-8'),$textFormat);
                $worksheet->writeString($i, $j++, html_entity_decode($users_result->email,ENT_QUOTES,'UTF-8'),$textFormat);                
                $worksheet->writeString($i, $j++, html_entity_decode($users_result->password,ENT_QUOTES,'UTF-8'),$textFormat);
                $worksheet->writeString($i, $j++, html_entity_decode($this->getStatus($users_result->status),ENT_QUOTES,'UTF-8'),$textFormat);
                $k++;
                $i += 1;
                $j = 0;
            }
        }
    }
}

?>