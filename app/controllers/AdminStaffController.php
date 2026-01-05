<?php
require_once __DIR__ . '/../models/StaffModel.php';

// IMPORT THƯ VIỆN EXCEL (Bắt buộc phải có đoạn này mới xuất được file)
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class AdminStaffController {
    private $staffModel;

    public function __construct() {
        $this->staffModel = new StaffModel();
    }

    // 1. Hiển thị danh sách nhân viên & Lịch làm hôm nay
    public function index() {
        $selectedDate = $_GET['date'] ?? date('Y-m-d');
        
        $staffs = $this->staffModel->getAllStaff();
        $dailyRoster = $this->staffModel->getDailyRoster($selectedDate);

        require __DIR__ . '/../views/staff/index.php';
    }

    // 2. Xử lý thêm lịch (Assign Shift)
    public function storeSchedule() {
        if ($_SERVER['REQUEST_METHOD'] === 'POST') {
            $user_id = $_POST['user_id'];
            $work_date = $_POST['work_date'];
            $shift_code = $_POST['shift_code']; // 1, 2, 3

            // Định nghĩa cứng các ca làm việc
            $shifts = [
                '1' => ['name' => 'Ca Sáng',  'start' => '07:00', 'end' => '12:00'],
                '2' => ['name' => 'Ca Chiều', 'start' => '12:00', 'end' => '17:00'],
                '3' => ['name' => 'Ca Tối',   'start' => '17:00', 'end' => '22:00'],
            ];

            if (isset($shifts[$shift_code])) {
                $shift = $shifts[$shift_code];
                // Gọi model thêm lịch
                $this->staffModel->assignShift($user_id, $work_date, $shift['name'], $shift['start'], $shift['end']);
            }
            
            header("Location: /GocCaPhe/public/index.php?url=admin/staff&date=$work_date");
            exit;
        }
    }

    // 3. Xóa lịch
    public function deleteSchedule() {
        $id = $_GET['id'];
        $this->staffModel->deleteShift($id);
        header('Location: ' . $_SERVER['HTTP_REFERER']);
        exit;
    }

    // --- CÁC CHỨC NĂNG MỚI BỔ SUNG ---

    // 4. Hiển thị Form Sửa (Edit)
    public function editSchedule() {
        $id = $_GET['id'] ?? null;
        if (!$id) {
            die("Không tìm thấy ID lịch làm việc.");
        }

        // Lấy thông tin ca làm việc hiện tại
        $shift = $this->staffModel->getShiftById($id);
        // Lấy danh sách nhân viên để chọn người thay thế
        $staffs = $this->staffModel->getAllStaff();

        // View sửa (Bạn cần tạo file này, xem code bên dưới)
        require __DIR__ . '/../views/staff/edit.php';
    }

    // --- CẬP NHẬT LOGIC UPDATE (ĐỔI NGƯỜI + ĐỔI CA) ---
    public function updateSchedule() {
        if ($_SERVER['REQUEST_METHOD'] === 'POST') {
            $schedule_id = $_POST['schedule_id'];
            $user_id = $_POST['user_id'];
            $current_date = $_POST['current_date'];
            $shift_code = $_POST['shift_code'];

            // Định nghĩa lại thông tin ca
            $shifts = [
                '1' => ['name' => 'Ca Sáng',  'start' => '07:00', 'end' => '12:00'],
                '2' => ['name' => 'Ca Chiều', 'start' => '12:00', 'end' => '17:00'],
                '3' => ['name' => 'Ca Tối',   'start' => '17:00', 'end' => '22:00'],
            ];

            if (isset($shifts[$shift_code])) {
                $s = $shifts[$shift_code];
                // Gọi hàm update mới trong Model
                $this->staffModel->updateShiftDetails($schedule_id, $user_id, $s['name'], $s['start'], $s['end']);
            }

            header("Location: /GocCaPhe/public/index.php?url=admin/staff&date=$current_date");
            exit;
        }
    }

    // --- LOGIC XUẤT EXCEL DẠNG BẢNG MA TRẬN (MỚI) ---
    public function exportWeekly() {
        $inputDate = $_GET['date'] ?? date('Y-m-d');
        $ts = strtotime($inputDate);
        
        // Xác định ngày đầu tuần (T2) và cuối tuần (CN)
        $startOfWeek = date('Y-m-d', strtotime('monday this week', $ts));
        $endOfWeek   = date('Y-m-d', strtotime('sunday this week', $ts));

        // Lấy dữ liệu thô
        $rawData = $this->staffModel->getWeeklyRoster($startOfWeek, $endOfWeek);

        // --- BƯỚC 1: CHUẨN BỊ DỮ LIỆU MA TRẬN ---
        // Cấu trúc: $matrix['Ca Sáng']['2023-10-23'] = "Tên NV"
        $matrix = [
            'Ca Sáng'  => [],
            'Ca Chiều' => [],
            'Ca Tối'   => []
        ];

        // Tạo mảng ngày trong tuần để làm Header
        $weekDates = [];
        for ($i = 0; $i < 7; $i++) {
            $d = date('Y-m-d', strtotime($startOfWeek . " +$i days"));
            $weekDates[] = $d;
        }

        // Đổ dữ liệu vào matrix
        foreach ($rawData as $row) {
            $sName = $row['shift_name'];
            $wDate = $row['work_date'];
            $uName = $row['staff_name'];

            if (isset($matrix[$sName])) {
                // Nếu 1 ca có nhiều người, nối tên bằng dấu phẩy
                if (isset($matrix[$sName][$wDate])) {
                    $matrix[$sName][$wDate] .= ", " . $uName;
                } else {
                    $matrix[$sName][$wDate] = $uName;
                }
            }
        }

        // --- BƯỚC 2: TẠO FILE EXCEL ---
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Tiêu đề
        $sheet->mergeCells('A1:H1');
        $sheet->setCellValue('A1', "LỊCH LÀM VIỆC TUẦN TỪ " . date('d/m', strtotime($startOfWeek)) . " ĐẾN " . date('d/m/Y', strtotime($endOfWeek)));
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Header Cột (Thứ 2 -> Chủ Nhật)
        $sheet->setCellValue('A3', 'CA LÀM VIỆC');
        $colIndex = 2; // Bắt đầu từ cột B
        $daysVi = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ Nhật'];

        foreach ($weekDates as $key => $date) {
            // Chuyển số thành chữ cái cột (0->B, 1->C...)
            $colString = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex);
            
            $headerText = $daysVi[$key] . "\n(" . date('d/m', strtotime($date)) . ")";
            $sheet->setCellValue($colString . '3', $headerText);
            $colIndex++;
        }

        // Style Header
        $sheet->getStyle('A3:H3')->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
        $sheet->getStyle('A3:H3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4CAF50');
        $sheet->getStyle('A3:H3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER)->setWrapText(true);
        $sheet->getRowDimension(3)->setRowHeight(40);

        // --- BƯỚC 3: IN DỮ LIỆU ---
        $rowNum = 4;
        $shiftTimeInfo = [
            'Ca Sáng' => "07:00 - 12:00",
            'Ca Chiều' => "12:00 - 17:00",
            'Ca Tối' => "17:00 - 22:00"
        ];

        foreach ($matrix as $shiftName => $datesData) {
            // Cột A: Tên ca + Giờ
            $cellContent = strtoupper($shiftName) . "\n" . $shiftTimeInfo[$shiftName];
            $sheet->setCellValue('A' . $rowNum, $cellContent);
            
            // Các cột ngày (B -> H)
            $colIndex = 2;
            foreach ($weekDates as $date) {
                $colString = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex);
                
                // Lấy tên nhân viên nếu có
                $staffNames = isset($datesData[$date]) ? $datesData[$date] : "";
                
                $sheet->setCellValue($colString . $rowNum, $staffNames);
                $colIndex++;
            }
            
            // Style cho dòng
            $sheet->getRowDimension($rowNum)->setRowHeight(60); // Cao hơn để chứa tên
            $rowNum++;
        }

        // --- BƯỚC 4: ĐỊNH DẠNG CHUNG ---
        // Kẻ khung
        $styleBorder = [
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, 
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true // Tự xuống dòng nếu tên dài
            ]
        ];
        $sheet->getStyle('A3:H6')->applyFromArray($styleBorder);

        // Độ rộng cột
        $sheet->getColumnDimension('A')->setWidth(20);
        for ($i = 2; $i <= 8; $i++) {
            $colString = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i);
            $sheet->getColumnDimension($colString)->setWidth(18);
        }

        // Xuất file
        $fileName = 'Lich_Tuan_' . date('W_Y', $ts) . '.xlsx';
        
        if (ob_get_contents()) ob_end_clean();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}