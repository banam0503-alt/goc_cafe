<?php
require_once __DIR__ . '/../models/Reservation.php';

class AdminReservationController {

    public function index() {
        if (session_status() == PHP_SESSION_NONE) {
            session_start();
        }
        
        $reservations = Reservation::all();
        // Kiểm tra lại đường dẫn file view của bạn, nếu đặt trong thư mục con thì sửa lại cho đúng
        require_once __DIR__ . '/../views/admin/reservations/index.php'; 
    }

    public function approve() {
        $id = $_GET['id'] ?? null;
        // Lấy ID nhân viên từ session
        $staffId = $_SESSION['user']['id'] ?? null; 

        if ($id && $staffId) {
            Reservation::approve($id, $staffId);
        }
        header("Location: index.php?url=admin/reservations");
        exit;
    }

    public function cancel() {
        $id = $_GET['id'] ?? null;
        $staffId = $_SESSION['user']['id'] ?? null;

        if ($id && $staffId) {
            Reservation::cancel($id, $staffId);
        }
        header("Location: index.php?url=admin/reservations");
        exit;
    }

    public function export() {
        if (session_status() == PHP_SESSION_NONE) session_start();

        // Lấy tên và vai trò người đang đăng nhập để ký tên
        $exportName = $_SESSION['user']['name'] ?? 'Admin';
        $exportRole = $_SESSION['user']['role'] ?? 'ADMIN';

        // Gọi model xuất file
        Reservation::exportAllReservationsExcel($exportName, $exportRole);
    }
}