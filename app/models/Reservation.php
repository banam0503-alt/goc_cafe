<?php
require_once __DIR__ . '/../../config/Database.php';

// Import đầy đủ các class cần thiết của PhpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class Reservation {

    private static function db() {
        return Database::connect();
    }

    /* ================= CREATE ================= */
    public static function create($data) {
        $conn = self::db();
        $sql = "INSERT INTO reservations (user_id, hoten, phone, songuoi, ngay, gio, ghichu, trangthai)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'CHO_DUYET')";
        $stmt = $conn->prepare($sql);
        $stmt->execute([
            $data['user_id'],
            $data['hoten'],
            $data['phone'],
            $data['songuoi'],
            $data['ngay'],
            $data['gio'],
            $data['ghichu']
        ]);
        // ĐÃ BỎ: self::appendExcel(...) -> Không dùng cách cũ nữa
    }

    /* ================= GET BY USER ================= */
    public static function getByUser($userId) {
        $conn = self::db();
        $stmt = $conn->prepare("SELECT * FROM reservations WHERE user_id = ? ORDER BY id DESC");
        $stmt->execute([$userId]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    /* ================= READ (ALL) ================= */
    public static function all() {
        $sql = "SELECT r.*, u.name AS staff_name, u.role AS staff_role 
                FROM reservations r
                LEFT JOIN users u ON r.approved_by = u.id
                ORDER BY r.id DESC";
        return self::db()->query($sql)->fetchAll(PDO::FETCH_ASSOC);
    }

    /* ================= APPROVE ================= */
    public static function approve($id, $staffId) {
        $sql = "UPDATE reservations SET trangthai='DA_DUYET', approved_by=?, approved_at=NOW() WHERE id=?";
        self::db()->prepare($sql)->execute([$staffId, $id]);
        // ĐÃ BỎ: self::updateExcelStatus(...)
    }

    /* ================= CANCEL ================= */
    public static function cancel($id, $staffId) {
        $sql = "UPDATE reservations SET trangthai='HUY', approved_by=?, approved_at=NOW() WHERE id=?";
        self::db()->prepare($sql)->execute([$staffId, $id]);
        // ĐÃ BỎ: self::updateExcelStatus(...)
    }

    /* ========================================================== */
    /* ================= HÀM XUẤT EXCEL MỚI ===================== */
    /* ========================================================== */
    public static function exportAllReservationsExcel($exportName, $exportRole)
    {
        // Xóa buffer
        if (ob_get_length()) ob_end_clean();

        // Lấy dữ liệu
        $data = self::all();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Danh sách đặt bàn');

        /* ================= 1. CHÈN LOGO ================= */
        $logoPath = __DIR__ . '/../../public/assets/img/logo1.jpg'; 
        if (file_exists($logoPath)) {
            $drawing = new Drawing();
            $drawing->setName('Logo');
            $drawing->setPath($logoPath);
            $drawing->setHeight(50);
            $drawing->setCoordinates('A1');
            $drawing->setWorksheet($sheet);
        }

        /* ================= 2. TIÊU ĐỀ QUÁN ================= */
        // Chỉ để tên quán và tiêu đề, bỏ ngày tháng ở đây
        $sheet->mergeCells('C4:I4'); 
        $sheet->setCellValue('C4', 'GÓC CAFE - DANH SÁCH ĐẶT BÀN');
        $sheet->getStyle('C4')->getFont()->setBold(true)->setSize(16)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color('FF7A4A2E'));
        $sheet->getStyle('C4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells('C5:I5');
        $sheet->setCellValue('C5', 'Hệ thống quản lý đặt chỗ');
        $sheet->getStyle('C5')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('C5')->getFont()->setItalic(true);

        /* ================= 3. HEADER BẢNG DỮ LIỆU ================= */
        $startRow = 7;
        $headers = [
            'A' => 'ID',
            'B' => 'Khách hàng',
            'C' => 'SĐT',
            'D' => 'Số người',
            'E' => 'Ngày đặt',
            'F' => 'Giờ',
            'G' => 'Ghi chú',
            'H' => 'Trạng thái',
            'I' => 'Người xử lý'
        ];

        foreach ($headers as $col => $text) {
            $cell = $col . $startRow;
            $sheet->setCellValue($cell, $text);
            
            // Style Header
            $sheet->getStyle($cell)->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
            $sheet->getStyle($cell)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('4CAF50'); // Màu xanh lá
            $sheet->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        }

        /* ================= 4. ĐỔ DỮ LIỆU ================= */
        $row = $startRow + 1;
        foreach ($data as $r) {
            $sheet->setCellValue("A$row", $r['id']);
            $sheet->setCellValue("B$row", $r['hoten']);
            $sheet->setCellValue("C$row", $r['phone']);
            $sheet->setCellValue("D$row", $r['songuoi']);
            $sheet->setCellValue("E$row", date('d/m/Y', strtotime($r['ngay'])));
            $sheet->setCellValue("F$row", $r['gio']);
            $sheet->setCellValue("G$row", $r['ghichu']);
            
            // Xử lý trạng thái
            $statusVi = 'Chờ duyệt';
            if ($r['trangthai'] == 'DA_DUYET') $statusVi = 'Đã duyệt';
            if ($r['trangthai'] == 'HUY') $statusVi = 'Đã hủy';
            $sheet->setCellValue("H$row", $statusVi);

            // Người xử lý
            $handler = $r['staff_name'] ? $r['staff_name'] . " (" . $r['staff_role'] . ")" : "";
            $sheet->setCellValue("I$row", $handler);

            // Kẻ bảng
            $sheet->getStyle("A$row:I$row")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            
            $row++;
        }

        // Auto width cột
        foreach (range('A', 'I') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        /* ================= 5. FOOTER (GIỐNG DOANH THU) ================= */
        $row += 3; // Cách ra 3 dòng

        // --- Dòng 1: Ngày tháng năm ---
        $sheet->mergeCells("G$row:I$row"); // Cột G đến I (Bên phải)
        $sheet->setCellValue("G$row", 'Hải Phòng, ngày ' . date('d') . ' tháng ' . date('m') . ' năm ' . date('Y'));
        // Tách biệt font và alignment để tránh lỗi
        $sheet->getStyle("G$row")->getFont()->setItalic(true);
        $sheet->getStyle("G$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // --- Dòng 2: Tiêu đề người ký ---
        $row++;
        $sheet->mergeCells("G$row:I$row");
        $sheet->setCellValue("G$row", 'QUẢN LÝ ĐẶT BÀN');
        $sheet->getStyle("G$row")->getFont()->setBold(true);
        $sheet->getStyle("G$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // --- Dòng 3: Tên người xuất file (Cách ra để ký) ---
        $row += 3; 
        $sheet->mergeCells("G$row:I$row");
        $sheet->setCellValue("G$row", "Người xuất file: $exportName ($exportRole)");
        $sheet->getStyle("G$row")->getFont()->setBold(true); // Nếu muốn nghiêng thì thêm ->setItalic(true) ở dòng riêng
        $sheet->getStyle("G$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        /* ================= 6. OUTPUT ================= */
        $filename = 'Danh_Sach_Dat_Ban_' . date('d-m-Y') . '.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}