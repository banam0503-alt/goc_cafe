<?php
class StaffModel {
    private $pdo;

    public function __construct() {
        $this->pdo = Database::connect();
    }

    // --- DÀNH CHO ADMIN ---

    // Lấy danh sách tất cả nhân viên
    public function getAllStaff() {
        $stmt = $this->pdo->query("SELECT * FROM users WHERE role = 'STAFF' ORDER BY id DESC");
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // Lấy lịch làm việc của TOÀN BỘ nhân viên trong 1 ngày cụ thể (để Admin xem hôm nay ai làm)
    public function getDailyRoster($date) {
        $sql = "SELECT s.*, u.name, u.avatar, u.phone 
                FROM work_schedules s
                JOIN users u ON s.user_id = u.id
                WHERE s.work_date = :date
                ORDER BY s.start_time ASC";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['date' => $date]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // Thêm lịch làm việc (Xếp lịch)
    public function assignShift($user_id, $date, $shift, $start, $end) {
        // Kiểm tra xem nhân viên đã có lịch vào ca đó ngày đó chưa để tránh trùng
        $check = $this->pdo->prepare("SELECT id FROM work_schedules WHERE user_id=? AND work_date=? AND shift_name=?");
        $check->execute([$user_id, $date, $shift]);
        if ($check->fetch()) return false; // Đã có lịch

        $sql = "INSERT INTO work_schedules (user_id, work_date, shift_name, start_time, end_time, hourly_rate) 
                VALUES (:uid, :date, :shift, :start, :end, 25000)";
        $stmt = $this->pdo->prepare($sql);
        return $stmt->execute([
            'uid' => $user_id,
            'date' => $date,
            'shift' => $shift,
            'start' => $start,
            'end' => $end
        ]);
    }

    // Xóa lịch làm việc
    public function deleteShift($id) {
        $stmt = $this->pdo->prepare("DELETE FROM work_schedules WHERE id = ?");
        return $stmt->execute([$id]);
    }

    // --- DÀNH CHO STAFF ---


    // Lấy lịch làm việc cá nhân (từ ngày hiện tại trở đi)
    public function getMySchedule($user_id) {
        $sql = "SELECT * FROM work_schedules 
                WHERE user_id = :uid AND work_date >= CURDATE()
                ORDER BY work_date ASC";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['uid' => $user_id]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }

    // --- BỔ SUNG CHO CHỨC NĂNG EDIT ---

    // Lấy thông tin chi tiết 1 ca làm việc theo ID
    public function getShiftById($id) {
        $stmt = $this->pdo->prepare("SELECT * FROM work_schedules WHERE id = ?");
        $stmt->execute([$id]);
        return $stmt->fetch(PDO::FETCH_ASSOC);
    }

    // Cập nhật nhân viên cho ca làm việc (Đổi người)
    public function updateShiftStaff($schedule_id, $new_user_id) {
        // Kiểm tra xem nhân viên mới đã có lịch trùng giờ ngày đó chưa (Optional logic)
        // Ở đây mình cập nhật thẳng user_id
        $stmt = $this->pdo->prepare("UPDATE work_schedules SET user_id = ? WHERE id = ?");
        return $stmt->execute([$new_user_id, $schedule_id]);
    }

    public function updateShiftDetails($id, $user_id, $shift_name, $start_time, $end_time) {
        $sql = "UPDATE work_schedules 
                SET user_id = :uid, shift_name = :sname, start_time = :start, end_time = :end 
                WHERE id = :id";
        $stmt = $this->pdo->prepare($sql);
        return $stmt->execute([
            'uid'   => $user_id,
            'sname' => $shift_name,
            'start' => $start_time,
            'end'   => $end_time,
            'id'    => $id
        ]);
    }

    // --- BỔ SUNG CHO CHỨC NĂNG EXPORT --- 

    // Lấy lịch làm việc trong khoảng thời gian (Từ ngày A đến ngày B)
    public function getWeeklyRoster($startDate, $endDate) {
        $sql = "SELECT s.work_date, s.shift_name, s.start_time, s.end_time, u.name as staff_name 
                FROM work_schedules s
                JOIN users u ON s.user_id = u.id
                WHERE s.work_date BETWEEN :start AND :end
                ORDER BY s.work_date ASC, s.start_time ASC";
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute(['start' => $startDate, 'end' => $endDate]);
        return $stmt->fetchAll(PDO::FETCH_ASSOC);
    }
}