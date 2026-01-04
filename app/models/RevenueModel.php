<?php
require_once __DIR__ . '/../../config/Database.php';

class RevenueModel {
    private $pdo;

    public function __construct() {
        $this->pdo = Database::connect();
    }

    // 1. Lấy thống kê theo NGÀY (Số cốc, Doanh thu)
    public function getDailyStats($date) {
        // Chỉ tính các đơn đã hoàn thành hoặc đã thanh toán
        $sql = "SELECT 
                    COUNT(DISTINCT o.id) as total_orders,
                    COALESCE(SUM(o.total_price), 0) as total_revenue,
                    COALESCE(SUM(oi.quantity), 0) as total_cups
                FROM orders o
                LEFT JOIN order_items oi ON o.id = oi.order_id
                WHERE DATE(o.created_at) = :date 
                AND o.status IN ('PAID', 'COMPLETED', 'APPROVED')"; // Trạng thái được tính là có tiền
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['date' => $date]);
        return $stmt->fetch(PDO::FETCH_ASSOC);
    }

    // 2. Tính lương nhân viên theo NGÀY (Dựa vào lịch làm việc)
    public function getDailyStaffCost($date) {
        // Công thức: (Giờ kết thúc - Giờ bắt đầu) * Lương theo giờ
        // TIMESTAMPDIFF(HOUR, start, end) tính số giờ
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

    // 3. Lấy thống kê theo THÁNG (Doanh thu & Lương)
    public function getMonthlyStats($month, $year) {
        // Tính tổng doanh thu tháng
        $sqlRevenue = "SELECT COALESCE(SUM(total_price), 0) as revenue 
                       FROM orders 
                       WHERE MONTH(created_at) = :m AND YEAR(created_at) = :y
                       AND status IN ('PAID', 'COMPLETED', 'APPROVED')";
        
        // Tính tổng lương tháng
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
    // 4. [MỚI] Lấy dữ liệu cho biểu đồ tròn (Doanh thu theo Danh mục trong tháng)
    public function getRevenueByCategory($month, $year) {
        $sql = "SELECT c.name, SUM(oi.quantity * oi.price) as total_money
                FROM order_items oi
                JOIN orders o ON oi.order_id = o.id
                JOIN products p ON oi.product_id = p.id
                JOIN categories c ON p.category_id = c.id
                WHERE MONTH(o.created_at) = :m 
                AND YEAR(o.created_at) = :y
                AND o.status IN ('PAID', 'COMPLETED', 'APPROVED')
                GROUP BY c.name";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['m' => $month, 'y' => $year]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }
}