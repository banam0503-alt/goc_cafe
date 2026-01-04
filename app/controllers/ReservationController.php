<?php
require_once __DIR__ . '/../models/Reservation.php';

class ReservationController {

    // 1. Hiển thị form đặt bàn
    public function create() {
        // Kiểm tra đăng nhập
        if (!isset($_SESSION['user'])) {
            header("Location: ?url=login");
            exit;
        }
        require_once __DIR__ . '/../views/user/create.php';
    }

    // 2. Xử lý lưu đơn đặt bàn
    public function store() {
        if (!isset($_SESSION['user'])) {
            header("Location: ?url=login");
            exit;
        }

        if ($_SERVER['REQUEST_METHOD'] === 'POST') {
            $data = [
                'user_id' => $_SESSION['user']['id'],
                'hoten'   => $_POST['hoten'],
                'phone'   => $_POST['phone'],
                'songuoi' => $_POST['songuoi'],
                'ngay'    => $_POST['ngay'],
                'gio'     => $_POST['gio'],
                'ghichu'  => $_POST['ghichu'] ?? ''
            ];

            Reservation::create($data);

            echo "<script>alert('Đặt bàn thành công! Vui lòng chờ xác nhận.'); window.location.href='?url=reservation/history';</script>";
        }
    }

    // 3. Xem lịch sử đặt bàn
    public function history() {
        if (!isset($_SESSION['user'])) {
            header("Location: ?url=login");
            exit;
        }

        $userId = $_SESSION['user']['id'];
        $myReservations = Reservation::getByUser($userId);

        require_once __DIR__ . '/../views/user/history.php';
    }
}