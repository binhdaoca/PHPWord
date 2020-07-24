<?php
require_once 'bootstrap.php';

// Tạo tài liệu mới ... 
$phpWord = new \PhpOffice\PhpWord\PhpWord();

$phpWord->setDefaultFontName('Times New Roman');
$phpWord->setDefaultFontSize(12);
/* Lưu ý: bất kỳ yếu tố nào bạn thêm vào tài liệu đều phải nằm trong Section. */

// Thêm một mục trống để tài liệu ...
$section = $phpWord->addSection();
// Thêm phần tử Văn bản vào Phần có phông chữ theo mặc định ... 

$section->addText(
    '"Learn from yesterday, live for today, hope for tomorrow. '
        . 'The important thing is not to stop questioning." '
        . '(Albert Einstein)'
);

/*
* Lưu ý: có thể tùy chỉnh kiểu phông chữ của thành phần Văn bản bạn thêm theo ba cách: 
* - nội tuyến; 
* - sử dụng kiểu phông chữ được đặt tên (đối tượng kiểu phông chữ mới sẽ được tạo ngầm); 
* - sử dụng đối tượng kiểu phông chữ được tạo rõ ràng.
 */

// Thêm phần tử Văn bản với phông chữ tùy chỉnh ... 
$section->addText(
    '"Great achievement is usually born of great sacrifice, '
        . 'and is never the result of selfishness." '
        . '(Napoleon Hill)',
    array('name' => 'Tahoma', 'size' => 10)
);

// Thêm phần tử văn bản với phông chữ được tùy chỉnh bằng cách sử dụng kiểu phông chữ có tên ... 
$fontStyleName = 'oneUserDefinedStyle';
$phpWord->addFontStyle(
    $fontStyleName,
    array('name' => 'Tahoma', 'size' => 10, 'color' => '1B2232', 'bold' => true)
);
$section->addText(
    '"The greatest accomplishment is not in never falling, '
        . 'but in rising again after you fall." '
        . '(Vince Lombardi)',
    $fontStyleName
);

// Thêm phần tử văn bản với phông chữ được tùy chỉnh bằng cách sử dụng đối tượng kiểu phông chữ được tạo rõ ràng ... 
$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setName('Tahoma');
$fontStyle->setSize(13);
$myTextElement = $section->addText('"Believe you can and you\'re halfway there." (Theodor Roosevelt)');
$myTextElement->setFontStyle($fontStyle);

// Lưu tài liệu dưới dạng tệp OOXML ... 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('helloWorld/helloWorld.docx');

// Lưu tài liệu dưới dạng tệp ODF ... 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
$objWriter->save('helloWorld/helloWorld.odt');

// Lưu tài liệu dưới dạng tệp HTML ... 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
$objWriter->save('helloWorld/helloWorld.html');

/* Lưu ý: chúng tôi bỏ qua RTF, vì nó không dựa trên XML và yêu cầu một ví dụ khác */
/* Lưu ý: chúng tôi bỏ qua PDF, vì phương pháp "HTML-to-PDF" được sử dụng để tạo tài liệu PDF. */