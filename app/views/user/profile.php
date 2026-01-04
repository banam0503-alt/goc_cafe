<?php require_once __DIR__ . '/../layouts/header.php'; ?>
<link rel="stylesheet" href="/GocCaPhe/public/assets/css/profile.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">

<div class="container mt-5">
    <div class="profile-wrapper">
        <aside class="profile-sidebar">
            <!-- SỬA LẠI ĐOẠN NÀY TRONG FILE VIEW (Xóa hết style inline đi) -->
            <div class="user-brief">
                <div class="sidebar-avatar-wrapper">
                    <img src="<?= $avatarPath ?>" alt="Avatar" id="sidebar-preview-img">
                    
                    <label for="file-input" class="upload-icon-label">
                        <i class="fas fa-plus-circle"></i>
                    </label>
                </div>
            </div>

            <ul class="sidebar-menu">
                <li class="active">
                    <a href="#"><i class="fas fa-user"></i> Hồ sơ của tôi</a>
                </li>
                <li>
                    <a href="?url=cart&status=PAID"><i class="fas fa-shopping-bag"></i> Đơn mua</a>
                </li>
                <li>
                    <a href="#"><i class="fas fa-bell"></i> Thông báo</a>
                </li>
            </ul>
        </aside>


        <main class="profile-main shadow-sm">
            <div class="profile-header">
                <h2>Hồ sơ của tôi</h2>
                <p>Quản lý thông tin hồ sơ để bảo mật tài khoản</p>
            </div>

            <div class="profile-body">
                <div class="profile-info">
                    <form action="/GocCaPhe/public/index.php?url=profile/update" method="POST" enctype="multipart/form-data">
                        <input type="file" name="avatar" id="file-input" hidden accept="image/*">

                        <div class="form-group-row">
                            <label>Tên đăng nhập</label>
                            <input type="text" name="name" value="<?= htmlspecialchars($user['name'] ?? '') ?>" required>
                        </div>

                        <div class="form-group-row">
                            <label>Email</label>
                            <input type="email" value="<?= htmlspecialchars($user['email'] ?? '') ?>" readonly class="readonly-input">
                        </div>

                        <div class="form-group-row">
                            <label>Số điện thoại</label>
                            <input type="text" name="phone" value="<?= htmlspecialchars($user['phone'] ?? '') ?>" placeholder="Nhập số điện thoại">
                        </div>

                        <div class="form-group-row">
                            <label>Địa chỉ</label>
                            <textarea name="address" rows="3" placeholder="Địa chỉ nhận hàng"><?= htmlspecialchars($user['address'] ?? '') ?></textarea>
                        </div>

                        <div class="form-group-row" style="margin-top: 30px;">
                            <label></label>
                            <button type="submit" class="btn-save-profile">Lưu thay đổi</button>
                        </div>
                    </form>

                    <hr class="my-5">

                    <h4 class="mb-4" style="color: #4b2e1f;">Đổi mật khẩu</h4>
                    <form action="/GocCaPhe/public/index.php?url=profile/update-password" method="POST">
                        <div class="form-group-row">
                            <label>Mật khẩu cũ</label>
                            <input type="password" name="old_password" required placeholder="********">
                        </div>
                        <div class="form-group-row">
                            <label>Mật khẩu mới</label>
                            <input type="password" name="new_password" required placeholder="Tối thiểu 6 ký tự">
                        </div>
                        <button type="submit" class="btn-save-profile" style="background: #28a745; color: #fff;">Cập nhật mật khẩu</button>
                    </form>
                </div>
            </div>
        </main>
    </div>
</div>

<script>
// Script xử lý xem trước ảnh ngay khi chọn file
document.getElementById('file-input').onchange = function (evt) {
    const [file] = this.files;
    if (file) {
        // Cập nhật ảnh preview ở sidebar
        document.getElementById('sidebar-preview-img').src = URL.createObjectURL(file);
    }
}
</script>

<?php require_once __DIR__ . '/../layouts/footer.php'; ?>