<?php
class UserModel {
    private $pdo;

    public function __construct() {
        // Sử dụng kết nối có sẵn từ class Database của dự án
        $this->pdo = Database::connect();
    }

    public function updateUserInfo($id, $name, $phone, $address, $avatar) {
        // Câu lệnh SQL cập nhật (không đụng đến email, password, role)
        $sql = "UPDATE users SET name = :name, phone = :phone, address = :address, avatar = :avatar WHERE id = :id";
        
        $stmt = $this->pdo->prepare($sql);
        
        return $stmt->execute([
            'name' => $name,
            'phone' => $phone,
            'address' => $address,
            'avatar' => $avatar,
            'id' => $id
        ]);
    }
}
?>