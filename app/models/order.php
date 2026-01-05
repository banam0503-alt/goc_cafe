<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;  

class Order
{
    private PDO $pdo;

    public function __construct(PDO $pdo)
    {
        $this->pdo = $pdo;
    }

    // Lấy tất cả đơn hàng
    public function getAllOrders()
    {
        $sql = "
            SELECT 
                o.id,
                u.name AS customer_name,
                o.total_price,
                o.status,
                s.name AS staff_name,
                o.approved_at,
                o.created_at,
                o.reject_reason
            FROM orders o
            LEFT JOIN users u ON o.user_id = u.id
            LEFT JOIN users s ON o.approved_by = s.id
            ORDER BY o.created_at DESC
        ";
        return $this->pdo->query($sql)->fetchAll(PDO::FETCH_ASSOC);
    }

    // Xuất Excel
    public function exportAllOrdersExcel($exportName, $exportRole)
    {
        // Xóa buffer
        if (ob_get_length()) {
            ob_end_clean();
        }

        $data = $this->getAllOrders();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Tất cả đơn hàng');

        /* ================= 1. LOGO (CĂN GIỮA) ================= */
        $logoPath = __DIR__ . '/../../public/assets/img/logo1.jpg';
        if (file_exists($logoPath)) {
            $drawing = new Drawing();
            $drawing->setName('Logo');
            $drawing->setPath($logoPath);
            $drawing->setHeight(50);
            
            // Đặt logo ở cột D (giữa của A-H) và chỉnh offset một chút để vào chính giữa
            $drawing->setCoordinates('E1'); 
            $drawing->setOffsetX(30); 
            
            $drawing->setWorksheet($sheet);
        }

        /* ================= 2. TIÊU ĐỀ (BỎ NGÀY THÁNG Ở ĐÂY) ================= */
        // Dòng 4: Tên quán
        $sheet->mergeCells('A4:H4');
        $sheet->setCellValue('A4', 'GÓC CAFE - QUẢN LÝ ĐƠN HÀNG');
        
        // Tách biệt Font và Alignment để tránh lỗi
        $sheet->getStyle('A4')->getFont()->setBold(true)->setSize(16)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color('FF7A4A2E'));
        $sheet->getStyle('A4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // (Đã xóa phần ngày tháng ở dòng 5 theo yêu cầu)

        /* ================= 3. HEADER TABLE ================= */
        $startRow = 7;

        $headers = [
            'A' => 'ID đơn',
            'B' => 'Khách hàng',
            'C' => 'Tổng tiền',
            'D' => 'Trạng thái',
            'E' => 'Người xử lý',
            'F' => 'Ngày duyệt',
            'G' => 'Ngày tạo',
            'H' => 'Lý do hủy'
        ];

        foreach ($headers as $col => $text) {
            $cell = $col . $startRow;
            $sheet->setCellValue($cell, $text);
            
            // Style Header
            $sheet->getStyle($cell)->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
            $sheet->getStyle($cell)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('4CAF50'); // Màu xanh lá đồng bộ
            $sheet->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        }

        /* ================= 4. DATA ================= */
        $row = $startRow + 1;
        foreach ($data as $o) {
            $sheet->setCellValue("A$row", $o['id']);
            $sheet->setCellValue("B$row", $o['customer_name'] ?? '—');
            $sheet->setCellValue("C$row", number_format($o['total_price'], 0, ',', '.') . ' ₫');
            $sheet->setCellValue("D$row", $o['status']);
            $sheet->setCellValue("E$row", $o['staff_name'] ?? '—');
            $sheet->setCellValue("F$row", $o['approved_at']);
            $sheet->setCellValue("G$row", $o['created_at']);
            $sheet->setCellValue("H$row", $o['reject_reason']);
            
            // Kẻ khung cho dòng dữ liệu
            $sheet->getStyle("A$row:H$row")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); // Căn giữa ID
            
            $row++;
        }

        // Auto width
        foreach (range('A', 'H') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        /* ================= 5. FOOTER (CHUYỂN NGÀY THÁNG XUỐNG DƯỚI) ================= */
        $row += 3; // Cách ra 3 dòng

        // --- Dòng 1: Ngày tháng năm (Cột F-H: Bên phải) ---
        $sheet->mergeCells("F$row:H$row");
        $sheet->setCellValue("F$row", 'Hải Phòng, ngày ' . date('d') . ' tháng ' . date('m') . ' năm ' . date('Y'));
        $sheet->getStyle("F$row")->getFont()->setItalic(true);
        $sheet->getStyle("F$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // --- Dòng 2: Tiêu đề người ký ---
        $row++;
        $sheet->mergeCells("F$row:H$row");
        $sheet->setCellValue("F$row", 'QUẢN LÝ CỬA HÀNG');
        $sheet->getStyle("F$row")->getFont()->setBold(true);
        $sheet->getStyle("F$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // --- Dòng 3: Khoảng trống để ký và Tên người xuất ---
        $row += 3;
        $sheet->mergeCells("F$row:H$row");
        $sheet->setCellValue("F$row", "Người xuất file: $exportName ($exportRole)");
        $sheet->getStyle("F$row")->getFont()->setBold(true); // Nếu muốn nghiêng: ->setItalic(true) ở dòng riêng
        $sheet->getStyle("F$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        /* ================= 6. OUTPUT ================= */
        $fileName = 'Tat_ca_don_hang_' . date('d-m-Y') . '.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
