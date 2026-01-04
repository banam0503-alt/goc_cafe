<?php

class AuthController {

    // ================= LOGIN =================
    public function login() {
        require_once __DIR__ . '/../views/auth/login.php';
    }

    public function handleLogin() {

        if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
            header('Location: /GocCaPhe/public/index.php?url=login');
            exit;
        }

        $email = trim($_POST['email'] ?? '');
        $password = trim($_POST['password'] ?? '');

        if ($email === '' || $password === '') {
            $_SESSION['error'] = 'Vui lòng nhập đầy đủ email và mật khẩu';
            header('Location: /GocCaPhe/public/index.php?url=login');
            exit;
        }

        $pdo = Database::connect();

        $stmt = $pdo->prepare("SELECT * FROM users WHERE email = :email LIMIT 1");
        $stmt->execute(['email' => $email]);
        $user = $stmt->fetch(PDO::FETCH_ASSOC);

        if (!$user || !password_verify($password, $user['password'])) {
            $_SESSION['error'] = 'Email hoặc mật khẩu không đúng';
            header('Location: /GocCaPhe/public/index.php?url=login');
            exit;
        }

        // 1. Xóa trường password hash để bảo mật (không lưu mật khẩu vào session)
        unset($user['password']);
        
        $_SESSION['user'] = $user;

        switch ($user['role']) {
            case 'ADMIN':
                header("Location: /GocCaPhe/public/index.php?url=admin");
                break;
            case 'STAFF':
                header("Location: /GocCaPhe/public/index.php?url=staff");
                break;
            default:
                header("Location: /GocCaPhe/public/index.php");
        }
        exit;
    }

    // ================= REGISTER =================

    // HIỂN THỊ FORM ĐĂNG KÝ
    public function register() {
        require_once __DIR__ . '/../views/auth/register.php';
    }

    public function handleRegister() {

    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        header('Location: /GocCaPhe/public/index.php?url=register');
        exit;
    }

    $name             = trim($_POST['name'] ?? '');
    $email            = trim($_POST['email'] ?? '');
    $password         = trim($_POST['password'] ?? '');
    $passwordConfirm  = trim($_POST['password_confirm'] ?? '');

    // 1. kiểm tra rỗng
    if ($name === '' || $email === '' || $password === '' || $passwordConfirm === '') {
        $_SESSION['error'] = 'Vui lòng nhập đầy đủ thông tin';
        header('Location: /GocCaPhe/public/index.php?url=register');
        exit;
    }

    // 2. kiểm tra nhập lại mật khẩu
    if ($password !== $passwordConfirm) {
        $_SESSION['error'] = 'Mật khẩu nhập lại không khớp';
        header('Location: /GocCaPhe/public/index.php?url=register');
        exit;
    }

    // 3. kiểm tra độ mạnh mật khẩu
    if (
        strlen($password) < 8 ||
        !preg_match('/[A-Z]/', $password) ||
        !preg_match('/[a-z]/', $password) ||
        !preg_match('/[0-9]/', $password)
    ) {
        $_SESSION['error'] =
            'Mật khẩu phải ≥ 8 ký tự, gồm chữ hoa, chữ thường và số';
        header('Location: /GocCaPhe/public/index.php?url=register');
        exit;
    }

    $pdo = Database::connect();

    // 4. kiểm tra email tồn tại
    $stmt = $pdo->prepare("SELECT id FROM users WHERE email = :email");
    $stmt->execute(['email' => $email]);
    if ($stmt->fetch()) {
        $_SESSION['error'] = 'Email đã tồn tại';
        header('Location: /GocCaPhe/public/index.php?url=register');
        exit;
    }

    // 5. hash mật khẩu
    $hashPassword = password_hash($password, PASSWORD_DEFAULT);

    // 6. insert user (CÓ name)
    $stmt = $pdo->prepare(
        "INSERT INTO users (name, email, password, role)
         VALUES (:name, :email, :password, 'USER')"
    );
    $stmt->execute([
        'name'     => $name,
        'email'    => $email,
        'password' => $hashPassword
    ]);

    // 7. quay về login
    $_SESSION['success'] = 'Đăng ký thành công, vui lòng đăng nhập';
    header('Location: /GocCaPhe/public/index.php?url=login');
    exit;
}


    // ================= LOGOUT =================
    public function logout() {
        session_destroy();
        header("Location: /GocCaPhe/public/index.php?url=login");
        exit;
    }
}
