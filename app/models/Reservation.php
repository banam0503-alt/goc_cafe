<?php
require_once __DIR__ . '/../../config/Database.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Reservation {

    private static function db() {
        return Database::connect();
    }

    private static function excelPath() {
        return __DIR__ . '/../../storage/reservations.xlsx';
    }

    /* ================= CREATE (Cập nhật thêm user_id) ================= */
    public static function create($data) {
        $conn = self::db();

        // Thêm user_id vào câu lệnh INSERT
        $sql = "INSERT INTO reservations
                (user_id, hoten, phone, songuoi, ngay, gio, ghichu, trangthai)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'CHO_DUYET')";

        $stmt = $conn->prepare($sql);
        $stmt->execute([
            $data['user_id'], // Thêm tham số này
            $data['hoten'],
            $data['phone'],
            $data['songuoi'],
            $data['ngay'],
            $data['gio'],
            $data['ghichu']
        ]);

        $id = $conn->lastInsertId();
        // Excel giữ nguyên logic cũ (có thể thêm cột User ID vào excel nếu cần)
        self::appendExcel($id, $data, 'CHO_DUYET');
    }

    /* ================= GET BY USER (MỚI - Xem lịch sử) ================= */
    public static function getByUser($userId) {
        $conn = self::db();
        $stmt = $conn->prepare("SELECT * FROM reservations WHERE user_id = ? ORDER BY id DESC");
        $stmt->execute([$userId]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    /* ================= READ (SỬA) ================= */
    public static function all() {
        // LEFT JOIN với bảng users để lấy tên và vai trò của người duyệt
        $sql = "SELECT r.*, u.name AS staff_name, u.role AS staff_role 
                FROM reservations r
                LEFT JOIN users u ON r.approved_by = u.id
                ORDER BY r.id DESC";
                
        return self::db()->query($sql)->fetchAll(PDO::FETCH_ASSOC);
    }

    /* ================= APPROVE (SỬA) ================= */
    public static function approve($id, $staffId) {
        $sql = "UPDATE reservations 
                SET trangthai='DA_DUYET', 
                    approved_by=?, 
                    approved_at=NOW() 
                WHERE id=?";
        self::db()->prepare($sql)->execute([$staffId, $id]);

        self::updateExcelStatus($id, 'DA_DUYET');
    }

    /* ================= CANCEL (SỬA) ================= */
    public static function cancel($id, $staffId) {
        $sql = "UPDATE reservations 
                SET trangthai='HUY', 
                    approved_by=?, 
                    approved_at=NOW() 
                WHERE id=?";
        self::db()->prepare($sql)->execute([$staffId, $id]);

        self::updateExcelStatus($id, 'HUY');
    }
    
    // Hàm thêm dòng mới vào Excel
    private static function appendExcel($id, $data, $status) {
        $path = self::excelPath();

        // Kiểm tra file có tồn tại không để load hoặc tạo mới
        if (file_exists($path) && filesize($path) > 0) {
            try {
                $spreadsheet = IOFactory::load($path);
                $sheet = $spreadsheet->getActiveSheet();
            } catch (Exception $e) {
                $spreadsheet = self::createNewSpreadsheet();
                $sheet = $spreadsheet->getActiveSheet();
            }
        } else {
            $spreadsheet = self::createNewSpreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
        }

        // Tìm dòng trống cuối cùng
        $row = $sheet->getHighestRow() + 1;

        // Ghi dữ liệu
        $sheet->fromArray([
            $id,
            $data['hoten'],
            $data['phone'],
            $data['songuoi'],
            $data['ngay'],
            $data['gio'],
            $data['ghichu'],
            $status // Cột H (Cột thứ 8)
        ], null, "A{$row}");

        // Lưu file
        $writer = new Xlsx($spreadsheet);
        $writer->save($path);
    }

    // Hàm cập nhật trạng thái (Duyệt/Hủy) trong Excel
    private static function updateExcelStatus($id, $status) {
        $path = self::excelPath();
        
        // Nếu file không tồn tại thì không làm gì cả
        if (!file_exists($path) || filesize($path) == 0) return;

        try {
            $spreadsheet = IOFactory::load($path);
            $sheet = $spreadsheet->getActiveSheet();
            $found = false;

            // Duyệt qua các dòng để tìm ID khớp
            // Bắt đầu từ dòng 2 (vì dòng 1 là tiêu đề)
            foreach ($sheet->getRowIterator(2) as $row) {
                $rowIndex = $row->getRowIndex();
                $cellId = $sheet->getCell("A{$rowIndex}")->getValue(); // Cột A là ID

                if ($cellId == $id) {
                    // Cập nhật cột H (Trạng thái)
                    $sheet->setCellValue("H{$rowIndex}", $status);
                    $found = true;
                    break; 
                }
            }

            // Nếu tìm thấy và sửa xong thì lưu lại
            if ($found) {
                $writer = new Xlsx($spreadsheet);
                $writer->save($path);
            }
        } catch (Exception $e) {
            // Ghi log lỗi nếu cần: error_log($e->getMessage());
        }
    }

    // Hàm tạo file Excel mới kèm tiêu đề
    private static function createNewSpreadsheet() {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        // Tạo tiêu đề cột
        $sheet->fromArray([
            ['ID', 'Họ tên', 'Phone', 'Số người', 'Ngày', 'Giờ', 'Ghi chú', 'Trạng thái']
        ]);
        return $spreadsheet;
    }
}