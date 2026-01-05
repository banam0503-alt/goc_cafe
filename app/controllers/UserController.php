<?php
require_once __DIR__ . '/../models/UserModel.php'; // Nhớ import Model

class UserController {
    
    public function profile() {
        // Kiểm tra đăng nhập
        if (!isset($_SESSION['user'])) {
            header('Location: /GocCaPhe/public/index.php?url=login');
            exit();
        }

        $user = $_SESSION['user'];
        // Xử lý đường dẫn avatar hiển thị (nếu null thì dùng ảnh mặc định)
        $avatarPath = !empty($user['avatar']) ? '/GocCaPhe/public/' . $user['avatar'] : '/GocCaPhe/public/assets/images/default-avatar.jng';
        
        require_once __DIR__ . '/../views/user/profile.php';
    }

    public function updateProfile() {
        if ($_SERVER['REQUEST_METHOD'] === 'POST') {
            // 1. Lấy dữ liệu từ Form
            $id = $_SESSION['user']['id'];
            $name = trim($_POST['name']);
            $phone = trim($_POST['phone']);
            $address = trim($_POST['address']);

            // --- THÊM CHECK SỐ ĐIỆN THOẠI TẠI ĐÂY ---
            // Regex: /^0\d{9}$/ 
            // ^0    : Bắt đầu bằng số 0
            // \d{9} : Theo sau là đúng 9 chữ số (tổng cộng là 10 số)
            // $     : Kết thúc chuỗi
            if (!preg_match('/^0\d{9}$/', $phone)) {
                // Nếu không đúng định dạng, quay lại trang profile và báo lỗi
                header('Location: /GocCaPhe/public/index.php?url=profile&status=error_phone');
                exit; // Dừng code ngay lập tức, không cho chạy tiếp xuống dưới
            }
            // ------------------------------------------
            
            // 2. Xử lý Avatar
            // Mặc định lấy avatar cũ từ session
            $avatarPath = $_SESSION['user']['avatar']; 

            // Kiểm tra xem người dùng có upload file không
            if (isset($_FILES['avatar']) && $_FILES['avatar']['error'] === 0) {
                $file = $_FILES['avatar'];
                $fileName = time() . '_' . basename($file['name']); // Đổi tên file để tránh trùng
                $targetDir = __DIR__ . '/../../public/assets/uploads/'; // Đường dẫn thư mục lưu ảnh
                $targetFile = $targetDir . $fileName;

                // Tạo thư mục nếu chưa tồn tại
                if (!is_dir($targetDir)) {
                    mkdir($targetDir, 0777, true);
                }

                // Di chuyển file và cập nhật đường dẫn mới
                if (move_uploaded_file($file['tmp_name'], $targetFile)) {
                    // Lưu đường dẫn tương đối vào DB
                    $avatarPath = 'assets/uploads/' . $fileName;
                }
            }

            // 3. Gọi Model để update vào Database
            $userModel = new UserModel();
            $result = $userModel->updateUserInfo($id, $name, $phone, $address, $avatarPath);

            if ($result) {
                // 4. CẬP NHẬT LẠI SESSION (Rất quan trọng)
                $_SESSION['user']['name'] = $name;
                $_SESSION['user']['phone'] = $phone;
                $_SESSION['user']['address'] = $address;
                $_SESSION['user']['avatar'] = $avatarPath;

                // Redirect về trang profile kèm thông báo thành công
                header('Location: /GocCaPhe/public/index.php?url=profile&status=success');
            } else {
                echo "Có lỗi xảy ra khi cập nhật.";
            }
        }
    }
    // ... (Giữ nguyên phần updatePassword cũ của bạn) ...
}
?>