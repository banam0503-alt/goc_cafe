<?php
$user = $_SESSION['user'] ?? null;
$role = $user['role'] ?? null;
$current_url = $_GET['url'] ?? 'index';
?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Góc Cà Phê</title>

    <!-- CSS GLOBAL -->
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/base.css">
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/header.css">
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/trangchu.css">
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/footer.css">
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/style.css">
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/introduce.css">
</head>
<body>


<header class="header">
    <div class="header-container">

        <!-- LOGO -->
        <div class="logo">
            <a href="/GocCaPhe/public/index.php">
                ☕ Góc Cà Phê
            </a>
        </div>

        <!-- MENU -->
        <nav class="nav-menu">
            <a href="/GocCaPhe/public/index.php" 
               class="<?= ($current_url == 'index') ? 'active' : '' ?>">Trang chủ</a>
            
            <a href="/GocCaPhe/public/index.php?url=menu" 
               class="<?= ($current_url == 'menu') ? 'active' : '' ?>">Sản phẩm</a>
            
            <a href="/GocCaPhe/public/index.php?url=booking" 
               class="<?= ($current_url == 'booking') ? 'active' : '' ?>">Đặt bàn</a>
            
            <a  href="/GocCaPhe/public/index.php?url=gioithieu" 
               class="<?= ($current_url == 'gioithieu') ? 'active' : '' ?>"> Giới thiệu </a>
        </nav>

        <!-- USER ACTION -->
        <div class="nav-user">

            <?php if ($role === 'USER'): ?>
                <a href="/GocCaPhe/public/index.php?url=cart" class="btn-cart">
                    🛒 Giỏ hàng
                </a>

            <?php elseif ($role === 'STAFF'): ?>
                <a href="/GocCaPhe/public/index.php?url=staff" class="btn-staff">
                    Nhân viên
                </a>

            <?php elseif ($role === 'ADMIN'): ?>
                <a href="/GocCaPhe/public/index.php?url=admin" class="btn-admin">
                    Admin Panel
                </a>
            <?php endif; ?>
              <?php if ($user): ?>
                <div class="user-dropdown">
                    <span class="user-name-display"><?= htmlspecialchars($user['name']) ?> ▾</span>
                    <div class="dropdown-menu">
                        <a href="#">Tài khoản</a>
                        <a href="/GocCaPhe/public/index.php?url=logout">Đăng xuất</a>
                    </div>
                </div>
            <?php else: ?>
                <a href="/GocCaPhe/public/index.php?url=login">Đăng nhập</a>
                <a href="/GocCaPhe/public/index.php?url=register" class="btn-register">
                    Đăng ký
                </a>
            <?php endif; ?>

        </div>

    </div>
    
</header>