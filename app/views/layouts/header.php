<?php
require_once __DIR__ . '/../../../core/bootstrap.php';
$user = $_SESSION['user'] ?? null;
$role = $user['role'] ?? null;
$current_url = $_GET['url'] ?? 'index';

$cartCount = 0;
if ($user && $role === 'USER') {
    $cartCount = CartController::getCartCount($user['id']);
}

$avatarPath = '/GocCaPhe/public/assets/images/default-avatar.png'; // Ảnh mặc định
if ($user && !empty($user['avatar'])) {
    $avatarPath = '/GocCaPhe/public/' . $user['avatar'];
}

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
    <link rel="stylesheet" href="/GocCaPhe/public/assets/css/profile.css">
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
            
            <a href="/GocCaPhe/public/index.php?url=datban" 
               class="<?= ($current_url == 'datban') ? 'active' : '' ?>">Đặt bàn</a>
            
            <a  href="/GocCaPhe/public/index.php?url=gioithieu" 
               class="<?= ($current_url == 'gioithieu') ? 'active' : '' ?>"> Giới thiệu </a>
        </nav>

        <!-- USER ACTION -->
        <div class="nav-user">

            <?php if ($role === 'USER'): ?>
                <a href="/GocCaPhe/public/index.php?url=cart" class="btn-cart">
                    🛒 Giỏ hàng
                    <?php if ($cartCount > 0): ?>
                       <span class="cart-count" id="cart-count"><?= $cartCount ?></span>
                    <?php endif; ?>
            </a>

            <?php elseif ($role === 'STAFF'): ?>
                <a href="/GocCaPhe/public/index.php?url=staff" class="btn-staff">
                    Nhân viên
                </a>

            <?php elseif ($role === 'ADMIN'): ?>
                    Admin Panel
                </a>
            <?php endif; ?>
              <?php if ($user): ?>
                
                <div class="user-dropdown">
                  <div style="display:flex; align-items:center; gap:8px;">
    <span class="user-name-display">
        <?= htmlspecialchars($user['name']) ?> ▾
    </span>

    <img 
        src="<?= $avatarPath ?>" 
        alt="Avatar"
        style="
            width: 44px;
            height: 44px;
            border-radius: 50%;
            object-fit: cover;
            display: block;
            cursor: pointer;
        "
    >
</div>

                    <div class="user-menu"> 
                       <a href="/GocCaPhe/public/index.php?url=profile">Tài khoản</a>
                    <a href="/GocCaPhe/public/index.php?url=logout">Đăng xuất</a> 
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
    <!-- Bootstrap CSS -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">

<!-- Optional JS (cho modal, dropdown, tooltip)
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script> -->


<script>
function refreshCartCount(){
    fetch('/GocCaPhe/public/index.php?url=cart/count')
        .then(res => res.json())
        .then(data => {
            const el = document.getElementById('cart-count');
            if(!el) return;

            if(data.count > 0){
                el.textContent = data.count;
                el.style.display = 'inline-block';
            } else {
                el.style.display = 'none';
            }
        });
}

// gọi ngay khi load
refreshCartCount();

// gọi lại sau mỗi 2s (hoặc sau AJAX add/delete)
setInterval(refreshCartCount, 2000);
</script>