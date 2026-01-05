<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

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
        if (ob_get_length()) {
            ob_end_clean();
        }

        $data = $this->getAllOrders();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Tất cả đơn hàng');

        /* ================= LOGO ================= */
        $logoPath = __DIR__ . '/../../public/assets/img/logo1.jpg';
        if (file_exists($logoPath)) {
            $drawing = new Drawing();
            $drawing->setName('Logo');
            $drawing->setPath($logoPath);
            $drawing->setHeight(80);
            $drawing->setCoordinates('A1');
            $drawing->setWorksheet($sheet);
        }

        /* ================= TIÊU ĐỀ ================= */
        $sheet->mergeCells('A4:H4');
        $sheet->setCellValue('A4', 'GÓC CAFE');
        $sheet->getStyle('A4')->getFont()->setBold(true)->setSize(16);
        $sheet->getStyle('A4')->getAlignment()->setHorizontal('center');

        $sheet->mergeCells('A5:H5');
        $sheet->setCellValue(
            'A5',
            'Hải Phòng, ngày ' . date('d') . ' tháng ' . date('m') . ' năm ' . date('Y')
        );
        $sheet->getStyle('A5')->getAlignment()->setHorizontal('center');
        $sheet->getStyle('A5')->getFont()->setItalic(true);

        /* ================= HEADER TABLE ================= */
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
            $sheet->getStyle($cell)->getFont()->setBold(true);
            $sheet->getStyle($cell)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB('E9ECEF');
        }

        /* ================= DATA ================= */
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
            $row++;
        }

        foreach (range('A', 'H') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        /* ================= FOOTER ================= */
        $footerRow = $row + 2;

        $sheet->mergeCells("A$footerRow:H$footerRow");
        $sheet->setCellValue("A$footerRow", 'QUẢN LÝ ĐƠN HÀNG');
        $sheet->getStyle("A$footerRow")->getFont()->setBold(true);
        $sheet->getStyle("A$footerRow")->getAlignment()->setHorizontal('center');

        $footerRow += 2;
        $sheet->mergeCells("F$footerRow:H$footerRow");
        $sheet->setCellValue(
            "F$footerRow",
            "Người xuất file: $exportName ($exportRole)"
        );
        $sheet->getStyle("F$footerRow")->getFont()->setBold(true)->setItalic(true);

        /* ================= OUTPUT ================= */
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="tat_ca_don_hang.xlsx"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
