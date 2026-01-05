<?php
require_once __DIR__ . '/../models/RevenueModel.php';

class AdminRevenueController {
    private $revenueModel;

    public function __construct() {
        $this->revenueModel = new RevenueModel();
    }

    public function index() {
        // 1. Lấy tham số filter (Fix lỗi lọc)
        $filterDate = isset($_GET['date']) ? $_GET['date'] : date('Y-m-d');
        $filterMonth = isset($_GET['month']) ? $_GET['month'] : date('m');
        $filterYear = isset($_GET['year']) ? $_GET['year'] : date('Y');

        // 2. Lấy số liệu thống kê
        $dailyStats = $this->revenueModel->getDailyStats($filterDate);
        $dailySalary = $this->revenueModel->getDailyStaffCost($filterDate);
        $dailyProfit = $dailyStats['total_revenue'] - $dailySalary;

        $monthlyStats = $this->revenueModel->getMonthlyStats($filterMonth, $filterYear);
        $monthlyProfit = $monthlyStats['revenue'] - $monthlyStats['salary'];

        // 3. [MỚI] Lấy dữ liệu biểu đồ
        $chartData = $this->revenueModel->getRevenueByCategory($filterMonth, $filterYear);

        // Chuẩn bị dữ liệu cho Chart.js (Chuyển sang JSON)
        $chartLabels = [];
        $chartValues = [];
        foreach ($chartData as $item) {
            $chartLabels[] = $item['name'];
            $chartValues[] = $item['total_money'];
        }

        require __DIR__ . '/../views/admin/revenue/index.php';
    }

    // --- HÀM MỚI: XUẤT EXCEL ---
    public function export() {
        if (session_status() == PHP_SESSION_NONE) session_start();

        // 1. Lấy tháng/năm từ URL
        $month = $_GET['month'] ?? date('m');
        $year  = $_GET['year'] ?? date('Y');
        
        // 2. Lấy thông tin người xuất
        $exportName = $_SESSION['user']['name'] ?? 'Admin';
        
        // 3. Gọi Model xử lý xuất file
        $this->revenueModel->exportRevenueExcel($month, $year, $exportName);
    }
}