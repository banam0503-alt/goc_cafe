<?php
require_once __DIR__ . '/../../config/Database.php';

// Import thư viện Excel
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Color;

class RevenueModel {
    private $pdo;

    public function __construct() {
        $this->pdo = Database::connect();
    }

    // 1. Lấy thống kê theo NGÀY (Số cốc, Doanh thu)
    public function getDailyStats($date) {
        $sqlOrder = "SELECT 
                        COUNT(id) as total_orders,
                        COALESCE(SUM(total_price), 0) as total_revenue
                     FROM orders 
                     WHERE DATE(created_at) = :date 
                     AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        
        $stmtOrder = $this->pdo->prepare($sqlOrder);
        $stmtOrder->execute(['date' => $date]);
        $orderStats = $stmtOrder->fetch(PDO::FETCH_ASSOC);

        $sqlCups = "SELECT 
                        COALESCE(SUM(oi.quantity), 0) as total_cups
                    FROM order_items oi
                    JOIN orders o ON oi.order_id = o.id
                    WHERE DATE(o.created_at) = :date 
                    AND o.status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        
        $stmtCups = $this->pdo->prepare($sqlCups);
        $stmtCups->execute(['date' => $date]);
        $cupStats = $stmtCups->fetch(PDO::FETCH_ASSOC);

        return [
            'total_orders'  => $orderStats['total_orders'],
            'total_revenue' => $orderStats['total_revenue'],
            'total_cups'    => $cupStats['total_cups']
        ];
    }

    // 2. Tính lương nhân viên theo NGÀY
    public function getDailyStaffCost($date) {
        $sql = "SELECT 
                    COALESCE(SUM(
                        (TIME_TO_SEC(end_time) - TIME_TO_SEC(start_time)) / 3600 * hourly_rate
                    ), 0) as total_salary
                FROM work_schedules
                WHERE work_date = :date";

        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['date' => $date]);
        $result = $stmt->fetch(PDO::FETCH_ASSOC);
        return $result['total_salary'];
    }

    // 3. Lấy thống kê theo THÁNG
    public function getMonthlyStats($month, $year) {
        $sqlRevenue = "SELECT COALESCE(SUM(total_price), 0) as revenue 
                       FROM orders 
                       WHERE MONTH(created_at) = :m AND YEAR(created_at) = :y
                       AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        
        $sqlSalary = "SELECT 
                        COALESCE(SUM(
                            (TIME_TO_SEC(end_time) - TIME_TO_SEC(start_time)) / 3600 * hourly_rate
                        ), 0) as salary
                      FROM work_schedules
                      WHERE MONTH(work_date) = :m AND YEAR(work_date) = :y";

        $stmtRev = $this->pdo->prepare($sqlRevenue);
        $stmtRev->execute(['m' => $month, 'y' => $year]);
        
        $stmtSal = $this->pdo->prepare($sqlSalary);
        $stmtSal->execute(['m' => $month, 'y' => $year]);

        return [
            'revenue' => $stmtRev->fetchColumn(),
            'salary' => $stmtSal->fetchColumn()
        ];
    }

    // 4. Biểu đồ tròn
    public function getRevenueByCategory($month, $year) {
        $sql = "SELECT c.name, SUM(oi.quantity * oi.price) as total_money
                FROM order_items oi
                JOIN orders o ON oi.order_id = o.id
                JOIN products p ON oi.product_id = p.id
                JOIN categories c ON p.category_id = c.id
                WHERE MONTH(o.created_at) = :m 
                AND YEAR(o.created_at) = :y
                AND o.status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')
                GROUP BY c.name";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // 5. Lấy dữ liệu chi tiết cho Excel
    public function getDailyBreakdown($month, $year) {
        $sqlRevenue = "SELECT 
                        DATE(created_at) as report_date,
                        COUNT(id) as total_orders,
                        COALESCE(SUM(total_price), 0) as total_revenue
                       FROM orders
                       WHERE MONTH(created_at) = :m AND YEAR(created_at) = :y
                       AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')
                       GROUP BY DATE(created_at)";
        
        $sqlCups = "SELECT 
                        DATE(o.created_at) as report_date,
                        COALESCE(SUM(oi.quantity), 0) as total_cups
                    FROM order_items oi
                    JOIN orders o ON oi.order_id = o.id
                    WHERE MONTH(o.created_at) = :m AND YEAR(o.created_at) = :y
                    AND o.status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')
                    GROUP BY DATE(o.created_at)";

        $stmtRev = $this->pdo->prepare($sqlRevenue);
        $stmtRev->execute(['m' => $month, 'y' => $year]);
        $revenues = $stmtRev->fetchAll(PDO::FETCH_ASSOC);

        $stmtCup = $this->pdo->prepare($sqlCups);
        $stmtCup->execute(['m' => $month, 'y' => $year]);
        $cups = $stmtCup->fetchAll(PDO::FETCH_KEY_PAIR);

        $finalData = [];
        foreach ($revenues as $row) {
            $d = $row['report_date'];
            $finalData[] = [
                'date' => $d,
                'orders' => $row['total_orders'],
                'revenue' => $row['total_revenue'],
                'cups' => $cups[$d] ?? 0
            ];
        }
        return $finalData;
    }

    // 6. XUẤT EXCEL (ĐÃ SỬA LỖI)
    public function exportRevenueExcel($month, $year, $exportName) {
        if (ob_get_length()) ob_end_clean();

        $data = $this->getDailyBreakdown($month, $year);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle("Doanh thu T$month-$year");

        /* --- 1. LOGO --- */
        $logoPath = __DIR__ . '/../../public/assets/img/logo1.jpg';
        if (file_exists($logoPath)) {
            $drawing = new Drawing();
            $drawing->setName('Logo');
            $drawing->setPath($logoPath);
            $drawing->setHeight(50);
            $drawing->setCoordinates('A1');
            $drawing->setWorksheet($sheet);
        }

        /* --- 2. TIÊU ĐỀ --- */
        $sheet->mergeCells('B4:E4');
        $sheet->setCellValue('B4', 'BÁO CÁO DOANH THU THÁNG ' . $month . '/' . $year);
        $sheet->getStyle('B4')->getFont()->setBold(true)->setSize(16)->setColor(new Color('FF7A4A2E'));
        $sheet->getStyle('B4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // === SỬA LỖI Ở ĐÂY ===
        $sheet->mergeCells('B5:E5');
        $sheet->setCellValue('B5', 'Góc Cà Phê - Hệ thống quản lý');
        // Tách dòng setItalic và setHorizontal ra
        $sheet->getStyle('B5')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B5')->getFont()->setItalic(true);
        // =====================

        /* --- 3. HEADER BẢNG --- */
        $startRow = 7;
        $headers = [
            'A' => 'Ngày',
            'B' => 'Số đơn hàng',
            'C' => 'Số cốc bán',
            'D' => 'Doanh thu (VNĐ)',
            'E' => 'Ghi chú'
        ];

        foreach ($headers as $col => $text) {
            $cell = $col . $startRow;
            $sheet->setCellValue($cell, $text);
            $sheet->getStyle($cell)->applyFromArray([
                'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
                'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => '4CAF50']],
                'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
                'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
            ]);
        }

        /* --- 4. ĐỔ DỮ LIỆU --- */
        $row = $startRow + 1;
        $totalRevenueMonth = 0;
        
        foreach ($data as $item) {
            $sheet->setCellValue("A$row", date('d/m/Y', strtotime($item['date'])));
            $sheet->setCellValue("B$row", $item['orders']);
            $sheet->setCellValue("C$row", $item['cups']);
            $sheet->setCellValue("D$row", $item['revenue']);
            
            $sheet->getStyle("D$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');
            
            $sheet->getStyle("A$row:E$row")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A$row:C$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            $totalRevenueMonth += $item['revenue'];
            $row++;
        }

        /* --- 5. TỔNG KẾT --- */
        $sheet->setCellValue("A$row", "TỔNG CỘNG THÁNG $month");
        $sheet->mergeCells("A$row:C$row");
        $sheet->setCellValue("D$row", $totalRevenueMonth);
        
        $sheet->getStyle("A$row:E$row")->applyFromArray([
            'font' => ['bold' => true, 'size' => 12, 'color' => ['rgb' => 'FF0000']],
            'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'FFFFCC']],
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
        ]);
        $sheet->getStyle("D$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');
        $sheet->getStyle("A$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        foreach (range('A', 'E') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        /* --- 6. FOOTER --- */
        $row += 3;

        // Ngày tháng năm
        $sheet->mergeCells("C$row:E$row");
        $sheet->setCellValue("C$row", 'Hải Phòng, ngày ' . date('d') . ' tháng ' . date('m') . ' năm ' . date('Y'));
        // Sửa lỗi ở Footer (tách biệt Font và Alignment)
        $sheet->getStyle("C$row")->getFont()->setItalic(true);
        $sheet->getStyle("C$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $row++;
        // Người lập báo cáo
        $sheet->mergeCells("C$row:E$row");
        $sheet->setCellValue("C$row", "Người lập báo cáo");
        $sheet->getStyle("C$row")->getFont()->setBold(true);
        $sheet->getStyle("C$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $row += 3; // Ký tên
        $sheet->mergeCells("C$row:E$row");
        $sheet->setCellValue("C$row", $exportName);
        $sheet->getStyle("C$row")->getFont()->setBold(true);
        $sheet->getStyle("C$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        /* --- 7. OUTPUT --- */
        $filename = "Bao_Cao_Doanh_Thu_T{$month}_{$year}.xlsx";
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}