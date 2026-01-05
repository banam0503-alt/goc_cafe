<?php
require_once __DIR__ . '/../models/Order.php';


class OrderController
{
public function exportExcel()
{
    // session_start() CHỈ GỌI 1 LẦN Ở index.php
    $exportName = $_SESSION['user']['name'] ?? 'Không xác định';
    $exportRole = $_SESSION['user']['role'] ?? '';

    $pdo = Database::connect();
    $order = new Order($pdo);
    $order->exportAllOrdersExcel($exportName, $exportRole);
}


    public function index()
    {
        $pdo = Database::connect();

        // SỬA QUERY: 
        // 1. LEFT JOIN thêm 1 lần nữa với bảng users (đặt alias là 'staff') để lấy thông tin người duyệt.
        // 2. Lấy tất cả trạng thái thay vì chỉ 'PAID'.
        $sql = "
            SELECT 
                o.*, 
                u.name AS customer_name,
                staff.name AS staff_name,
                staff.role AS staff_role
            FROM orders o
            JOIN users u ON o.user_id = u.id
            LEFT JOIN users staff ON o.approved_by = staff.id
            WHERE o.status IN ('PAID', 'APPROVED', 'CANCELLED', 'SHIPPING', 'COMPLETED')
            ORDER BY o.created_at DESC
        ";

        $stmt = $pdo->query($sql);
        $orders = $stmt->fetchAll();

        require __DIR__ . '/../views/admin/orders/index.php';
    }

    public function approve()
    {
        $orderId = $_POST['id']; 
        $staffId = $_SESSION['user']['id'];

        $pdo = Database::connect();

        $stmt = $pdo->prepare("
            UPDATE orders
            SET status = 'APPROVED',
                approved_by = ?,
                approved_at = NOW() 
            WHERE id = ?
        ");

        $stmt->execute([$staffId, $orderId]);

        header('Location: ?url=admin/orders');
exit;
    }

    public function reject()
    {
        $orderId = $_POST['id']; 
        $reason = $_POST['reject_reason'] ?? '';
        $staffId = $_SESSION['user']['id'];

        $pdo = Database::connect();

        $stmt = $pdo->prepare("
            UPDATE orders
            SET status = 'CANCELLED',
                approved_by = ?,
                approved_at = NOW(),
                reject_reason = ?
            WHERE id = ?
        ");

        $stmt->execute([$staffId, $reason, $orderId]);

        header('Location: ?url=admin/orders');
exit;
    }
}
