<?php
require_once __DIR__ . '/../../config/Database.php';

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

    // 1. Thống kê NGÀY
    public function getDailyStats($date) {
        $sqlOrder = "SELECT COUNT(id) as total_orders, COALESCE(SUM(total_price), 0) as total_revenue
                     FROM orders WHERE DATE(created_at) = :date AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        $stmtOrder = $this->pdo->prepare($sqlOrder);
        $stmtOrder->execute(['date' => $date]);
        $orderStats = $stmtOrder->fetch(PDO::FETCH_ASSOC);

        $sqlCups = "SELECT COALESCE(SUM(oi.quantity), 0) as total_cups
                    FROM order_items oi JOIN orders o ON oi.order_id = o.id
                    WHERE DATE(o.created_at) = :date AND o.status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        $stmtCups = $this->pdo->prepare($sqlCups);
        $stmtCups->execute(['date' => $date]);
        $cupStats = $stmtCups->fetch(PDO::FETCH_ASSOC);

        return [
            'total_orders'  => $orderStats['total_orders'],
            'total_revenue' => $orderStats['total_revenue'],
            'total_cups'    => $cupStats['total_cups']
        ];
    }

    // 2. Thống kê THÁNG (Doanh thu)
    public function getMonthlyStats($month, $year) {
        $sql = "SELECT COALESCE(SUM(total_price), 0) as revenue 
                FROM orders 
                WHERE MONTH(created_at) = :m AND YEAR(created_at) = :y
                AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        return ['revenue' => $stmt->fetchColumn()];
    }

    // 3. Biểu đồ tròn
    public function getRevenueByCategory($month, $year) {
        $sql = "SELECT c.name, SUM(oi.quantity * oi.price) as total_money
                FROM order_items oi
                JOIN orders o ON oi.order_id = o.id
                JOIN products p ON oi.product_id = p.id
                JOIN categories c ON p.category_id = c.id
                WHERE MONTH(o.created_at) = :m AND YEAR(o.created_at) = :y
                AND o.status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')
                GROUP BY c.name";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // 4. [MỚI] Lấy chi phí tháng
    public function getMonthlyExpenses($month, $year) {
        $sql = "SELECT * FROM monthly_expenses WHERE month = :m AND year = :y";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        $result = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if (!$result) {
            return ['staff_cost' => 0, 'electricity_water_cost' => 0, 'materials_cost' => 0];
        }
        return $result;
    }

    // 5. [MỚI] Lưu chi phí tháng
    public function saveMonthlyExpenses($month, $year, $staff, $electric, $materials) {
        // Chúng ta đặt tên tham số khác nhau cho phần INSERT và UPDATE để tránh lỗi PDO
        $sql = "INSERT INTO monthly_expenses (month, year, staff_cost, electricity_water_cost, materials_cost) 
                VALUES (:m, :y, :s, :e, :mat)
                ON DUPLICATE KEY UPDATE 
                staff_cost = :s_update, 
                electricity_water_cost = :e_update, 
                materials_cost = :mat_update";
        
        $stmt = $this->pdo->prepare($sql);
        
        // Truyền đủ 8 tham số cho 8 vị trí giữ chỗ trong câu SQL
        return $stmt->execute([
            'm'   => $month, 
            'y'   => $year, 
            's'   => $staff, 
            'e'   => $electric, 
            'mat' => $materials,
            // Truyền lại giá trị cho phần Update
            's_update'   => $staff,
            'e_update'   => $electric,
            'mat_update' => $materials
        ]);
    }

    // 6. Dữ liệu chi tiết cho Excel
    public function getDailyBreakdown($month, $year) {
        $sql = "SELECT DATE(created_at) as report_date, COUNT(id) as total_orders, COALESCE(SUM(total_price), 0) as total_revenue
                FROM orders
                WHERE MONTH(created_at) = :m AND YEAR(created_at) = :y
                AND status IN ('PAID', 'APPROVED', 'SHIPPING', 'COMPLETED')
                GROUP BY DATE(created_at)";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // 7. [CẬP NHẬT] Xuất Excel
    public function exportRevenueExcel($month, $year, $exportName) {
        if (ob_get_length()) ob_end_clean();

        $data = $this->getDailyBreakdown($month, $year);
        $expenses = $this->getMonthlyExpenses($month, $year); // Lấy chi phí

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle("Loi_Nhuan_T$month-$year");

        // --- Header ---
        $sheet->mergeCells('A1:E1');
        $sheet->setCellValue('A1', 'BÁO CÁO LỢI NHUẬN THÁNG ' . $month . '/' . $year);
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(16);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // --- Table Header ---
        $startRow = 4;
        $sheet->setCellValue("A$startRow", "Ngày");
        $sheet->setCellValue("B$startRow", "Số đơn");
        $sheet->setCellValue("C$startRow", "Doanh thu (VNĐ)");
        $sheet->getStyle("A$startRow:C$startRow")->getFont()->setBold(true);

        // --- Data Loop ---
        $row = $startRow + 1;
        $totalRevenue = 0;
        foreach ($data as $item) {
            $sheet->setCellValue("A$row", date('d/m/Y', strtotime($item['report_date'])));
            $sheet->setCellValue("B$row", $item['total_orders']);
            $sheet->setCellValue("C$row", $item['total_revenue']);
            $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0');
            $totalRevenue += $item['total_revenue'];
            $row++;
        }

        // --- TÍNH TOÁN LỢI NHUẬN ---
        $row++;
        $sheet->setCellValue("B$row", "TỔNG DOANH THU:");
        $sheet->setCellValue("C$row", $totalRevenue);
        $sheet->getStyle("B$row:C$row")->getFont()->setBold(true);
        $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');
        
        $row++;
        $sheet->setCellValue("B$row", "(-) Tiền nhân viên:");
        $sheet->setCellValue("C$row", $expenses['staff_cost']);
        $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');

        $row++;
        $sheet->setCellValue("B$row", "(-) Tiền điện/nước:");
        $sheet->setCellValue("C$row", $expenses['electricity_water_cost']);
        $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');

        $row++;
        $sheet->setCellValue("B$row", "(-) Tiền nguyên liệu:");
        $sheet->setCellValue("C$row", $expenses['materials_cost']);
        $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');
        $sheet->getStyle("C$row")->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);

        $row++;
        $profit = $totalRevenue - ($expenses['staff_cost'] + $expenses['electricity_water_cost'] + $expenses['materials_cost']);
        $sheet->setCellValue("B$row", "LỢI NHUẬN THỰC TẾ:");
        $sheet->setCellValue("C$row", $profit);
        $sheet->getStyle("B$row:C$row")->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle("B$row:C$row")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFCC');
        $sheet->getStyle("C$row")->getNumberFormat()->setFormatCode('#,##0 "₫"');

        // --- Footer ---
        $row += 3;
        $sheet->setCellValue("B$row", "Người lập: " . $exportName);
        
        // Auto width
        foreach(range('A','C') as $col) $sheet->getColumnDimension($col)->setAutoSize(true);

        // Output
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="Bao_Cao_Loi_Nhuan_T'.$month.'_'.$year.'.xlsx"');
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}