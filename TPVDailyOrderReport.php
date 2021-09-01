<?php


/*
- Need to be done at the end of day at 11pm
- Get COMPLETED/CLOSED cases that was done during the appointment day 
- Not include CANCELLED cases 
- CLOSED DATE (NOT APPOINTMENT DATE)


REQUIRED DATE: 
*/


require 'vendor/autoload.php';
######PHPSpreadsheet######
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
date_default_timezone_set('America/Los_Angeles');


/////////////////////////////////////////////////////////////////////////////////////////////////
######CONNECT TO DATABASE#########

$server = "10.10.99.xx";
$username = "root";
$password = "xxx";
$database = "xxx";
//global $conn;
$conn = mysqli_connect($server, $username, $password, $database);
if(mysqli_connect_error())
{
    echo "Failed to connect: " .mysqli_connect_error();
    exit();
}


////////////////////////////////////////////////////////////////////////////////////////////////


chdir("Daily_reports");  

$spreadsheet = new Spreadsheet();
$writer = new Xlsx($spreadsheet);


$sheetIndex = $spreadsheet->getIndex($spreadsheet->getSheetByName('Worksheet'));
$spreadsheet->removeSheetByIndex($sheetIndex);
$myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'EPI');
$spreadsheet->addSheet($myWorkSheet,0);


$spreadsheet->getSheetByName('EPI');
$spreadsheet->setActiveSheetIndexByName('EPI');


////////SETTING THE TITLE ROW////////////

$active_sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);
$spreadsheet->getActiveSheet()->getStyle('A1:T500')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadsheet->getActiveSheet()->getStyle('O2:O500')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);//NO DECIMAL NUMBERS
$spreadsheet->getActiveSheet()->getStyle('Q2:Q500')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);//NO DECIMAL NUMBERS


$active_sheet->setCellValue('A1', 'Reference No');
$active_sheet->setCellValue('B1', 'Status');
$active_sheet->setCellValue('C1', 'SubStatus');
$active_sheet->setCellValue('D1', 'Model');
$active_sheet->setCellValue('E1', 'Serial Number');
$active_sheet->setCellValue('F1', 'Creation Date');
$active_sheet->setCellValue('G1', 'Closed Date');
$active_sheet->setCellValue('H1', 'Location Name');
$active_sheet->setCellValue('I1', 'Address 1' );
$active_sheet->setCellValue('J1', 'City' );
$active_sheet->setCellValue('K1', 'State' );
$active_sheet->setCellValue('L1', 'Zipcode' );
$active_sheet->setCellValue('M1', 'Phone' );
$active_sheet->setCellValue('N1', 'Complaint' );
$active_sheet->setCellValue('O1', 'Resolution');
$active_sheet->setCellValue('P1', 'Tier-2 tech support from ITI');
$active_sheet->setCellValue('Q1', 'Parts Needed');
$active_sheet->setCellValue('R1', 'Repair Action');
$active_sheet->setCellValue('S1', 'Part Replaced');
$active_sheet->setCellValue('T1', 'Part Shipped');
$active_sheet->setCellValue('U1', 'Part Shipped Track');
$active_sheet->setCellValue('V1', 'Part Shipped Date');
$active_sheet->setCellValue('W1', 'Part Returned Track');
$active_sheet->setCellValue('X1', 'Part Returned Date');
$active_sheet->setCellValue('Y1', 'Purchase Location');
$active_sheet->setCellValue('Z1', 'POP Date');

//////////////////////////////////////////////////////////////////////////////////////////////


/////////////////////// QUERY //////////////////////////////

//$reportDate = '2021-02-05';
$reportDate = date("Y-m-d");

$sql = "SELECT * FROM orders WHERE companyId IN (7,12) AND OEM = 'EPI' AND (status = 'WORKORDERSTATUS_COMPLETED' OR status ='WORKORDERSTATUS_CLOSED' OR status = 'WORKORDERSTATUS_CANCELLED') 
AND CAST(closedDate AS DATE)='$reportDate'";
$query = mysqli_query($conn,$sql);
$row = mysqli_num_rows($query);
echo $row;

for ($i=2; $i<=$row+1; $i++){
    $data =  mysqli_fetch_assoc($query);

    $referenceNo = $data['externalId'];
    echo "Ref No: $referenceNo ";
    $status = $data['status'];
    $substatus = $data['substatus'];
    $model = $data['model'];
    $serialNumber = $data['serialNumber'];
    $creationDate = substr($data['creationDate'],0,10);  
    $closeDate =substr($data['closedDate'],0,10);
    $locationName = $data['locationName'];
    $address1 = $data['street1'];
    $city = $data['city'];
    $state = $data['state'];
    $zipcode = $data['zipcode'];
    $phone = $data['phone'];
    $complaint = $data['problemDesc'];
    $resolution = $data['resolution'];
    $techSupport = substr($data['noteForTech'],23);
    $repairAction = $data['repairAction'];    
    $partReplaced = $data['partReplaced'];
    $purchaseLocation = $data['purchaseLocation'];
    $POPDate = substr($data['popDate'],0,10);

//////////////////GET PART SHIPPED//////////////////////////////////////////


    $sql1 = "SELECT * FROM part_repair WHERE externalId = '$referenceNo'";
    $query1 = mysqli_query($conn, $sql1);
   // while ($row1 = mysqli_fetch_array($query5, MYSQLI_ASSOC)){
    $row1 = mysqli_num_rows($query1);
    $partShipped = "";
    for($j=1; $j<=$row1; $j++){
        $data1 = mysqli_fetch_assoc($query1);
        $partShippedTrack = $data1['outBoundTracking'];
        $partShippedDate = $data1['partShippedDate'];
        $partReturnedTrack = $data1['returnTracking'];
        $partReturnedDate = $data1['partReturnedDate'];
      
        $partShipped .= $data1['partNumber']. ", ";
        //echo $partShipped;
           
        


    }

    echo $partShipped;


    



     ////////////////GET STATUS AND SUB STATUS DESC/////////////////
    $sql2 = "SELECT * FROM status where code = '$status'";
    $query2 = mysqli_query($conn, $sql2);
    while ($row2 = mysqli_fetch_array($query2, MYSQLI_ASSOC)){
        global $statusDescription;
        $statusDescription = $row2['description'];
    }

    $sql3 = "SELECT * FROM substatus where code = '$substatus'";
    $query3 = mysqli_query($conn, $sql3);
    while ($row3 = mysqli_fetch_array($query3, MYSQLI_ASSOC)){
        global $subDescription;
        $subDescription = $row3['description'];
    }


    /////////////Check repair action ///////////////////////
    /*
    if(strlen($partReplaced) >3){
        $repairAction = "Replaced";
    
    }else{
        $repairAction = "N/A";
    }
    */

    //global $partShipped;
    //echo $partShipped;
    $partNeeded = $partShipped;
    global $partShippedTrack;
    global $partShippedDate;
    global $partReturnedTrack;
    global $partReturnedDate;

    ///////////////////////MAP DATA FIELD///////////////////////////////////////

    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (1,$i, $referenceNo);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (2,$i, $statusDescription);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (3,$i, $subDescription);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (4,$i, $model);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (5,$i, $serialNumber);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (6,$i, $creationDate);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (7,$i, $closeDate);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (8,$i, $locationName);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (9,$i, $address1);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (10,$i, $city);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (11,$i, $state);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (12,$i, $zipcode);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (13,$i, $phone);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (14,$i, $complaint);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (15,$i, $resolution);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (16,$i, $techSupport);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (17,$i, $partNeeded);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (18,$i, $repairAction);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (19,$i, $partReplaced);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (20,$i, $partShipped);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (21,$i, $partShippedTrack);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (22,$i, $partShippedDate);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (23,$i, $partReturnedTrack);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (24,$i, $partReturnedDate);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (25,$i, $purchaseLocation);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow (26,$i, $POPDate);






    
//////////////////Save spreadsheet//////////////////////////

$writer->save("$reportDate TPV Daily Report.xlsx");
$attachment = "$reportDate TPV Daily Report.xlsx";


}

/////////////////////// SEND EMAIL //////////////////////////////



$email = "nguyenthaohien1997@gmail.com";

$to = $email;
$from = "Support@itiworldwide.com";
$fromName = "ITI Support";
//$subject = "Daily Work Order";
$subject = "$reportDate Daily Work Order";
$file = $attachment;

echo $file;
$htmlContent = '<h3> Daily Work Order Report generated by ITI system </h3>
                <p> This email is sent with an attachment. </p>';


///////////HEADER/////////////////////////
$header = "From: $fromName"." <".$from.">";


//////////BOUNDARY////////////////////////
$semi_rand = md5(time());
$mime_boundary = "==Multipart_Boundary_x{$semi_rand}x";

////////HEADER FOR ATTACHMENT/////////////
$header .= "\nMIME-Version: 1.0\n" . "Content-Type: multipart/mixed;\n" . " boundary=\"{$mime_boundary}\"";


//////////MULTIPART BOUNDARY///////////////

$message = "--{$mime_boundary}\n" . "Content-Type: text/html; charset=\"UTF-8\"\n" . 
           "Comtent-Transfer-Encoding: 7bit\n\n" . $htmlContent . "\n\n";

if(!empty($file) >0){
    if(is_file($file)){
        $message .= "--{$mime_boundary}\n";
        $fp = @fopen($file,"rb");
        //echo $fp;
        $data = @fread($fp, filesize($file));

        @fclose($fp);
        $data = chunk_split(base64_encode($data));
        $message .= "Content-Type: application/ontet-stream; name=\"".basename($file)."\"\n" . 
                    "Content-Description: ".basename($file). "\n" . 
                    "Content-Disposition: attachment;\n" . " filename=\"".basename($file)."\"; size=".filesize($file).";\n" . 
                    "Content-Transfer-Encoding: base64\n\n" . $data . "\n\n";

    }
}
$message .= "--{$mime_boundary}--";
$returnpath ="-f" . $from;

///////////////// Send email////////////////////////
$mail = @mail($to, $subject, $message, $header, $returnpath);
echo $mail?"<h1>Email Sent Successfully!</h1>":"<h1>Email sending failed.</h1>"; 

################################
?>

















