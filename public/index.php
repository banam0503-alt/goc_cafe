<?php session_start(); ?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="../assets/css/header.css">
    <link rel="stylesheet" href="../assets/css/trangchu.css">
    <title>Góc Cafe - Trang chủ</title>
</head>
<body>

<div class="nav">
    <a href="../pages/index.php">Trang chủ</a>
    <a href="../pages/menu.php">Menu</a>
    <a href="../pages/gioithieu.php">Giới thiệu</a>
    <a href="../pages/datban.php">Đặt bàn</a>

    <?php if(isset($_SESSION['user'])): ?>
        <span style="margin-left:20px;">
            Xin chào <?= htmlspecialchars($_SESSION['user']) ?>
        </span>
        <a href="logout.php">Đăng xuất</a>
    <?php else: ?>
        <a href="login.php">Đăng nhập</a>
        <a href="register.php">Đăng ký</a>
    <?php endif; ?>
</div>
<div class="banner-top">
  <div class="banner-title">
        <h1>HƯƠNG VỊ CÀ PHÊ TRUYỀN THỐNG</h1>
        <p>Nguyên chất – Chuẩn vị Việt Nam</p>
    </div>
    <img src="../assets/img/goc12.jpg" class="banner-img">
    <img src="../assets/img/gg.jpg" class="banner-img">
    <img src="../assets/img/goc.jpg" class="banner-img">

    <div class="menu-buttons">
        <a href="menu.php">MENU</a>
        <a href="datban.php">Đặt bàn</a>
        <a href="gioithieu.php">Giới thiệu</a>
    </div>
</div>
<!-- =================== HÌNH ẢNH DƯỚI BANNER =================== -->
<div class="sub-images">
  <img src="../assets/img/g1.jpg" alt="Ảnh 1" />
  <img src="../assets/img/g2.jpg" alt="Ảnh 2" />
  <img src="../assets/img/g3.jpg" alt="Ảnh 3" />
  <img src="../assets/img/g4.jpg" alt="Ảnh 4" />
</div>

<!-- =================== MỤC LỤC TRUNG TÂM =================== -->

      <div class="contact-container">
        <div class="contact-info">
          <h2>Liên hệ</h2>

          <p>
            <b>Địa chỉ:</b> Căn liền kề LK1-27 Chung cư Hoàng Huy, An Đồng, An
            Dương, Hải Phòng
          </p>
          <p><b>Hotline:</b> 0839331102</p>
          <p><b>Email:</b> banam0503@gmail.com</p>

          <p>
            📍
            <a href="https://maps.app.goo.gl/NV9a5nx6WNSSrYyn9" target="_blank">
              Xem vị trí trên Google Maps
            </a>
          </p>
        </div>

        <div class="contact-map">
          <iframe
            src="https://maps.google.com/maps?q=G%C3%B3c%20-%20Tr%C3%A0%20%26%20Cafe,%20An%20D%C6%B0%C6%A1ng,%20H%E1%BA%A3i%20Ph%C3%B2ng&t=&z=15&ie=UTF8&iwloc=&output=embed"
          >
          </iframe>
        </div>
      </div>
    </section>

    <!-- Footer -->
    <footer class="footer">
      © 2025 Quán Cafe Góc Cafe. Mọi quyền được bảo lưu.
    </footer>

</body>
</html>
