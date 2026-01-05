<?php require_once __DIR__ . '/../layouts/header.php'; ?>

<div class="container mt-5 mb-5">
    <div class="text-center mb-4">
        <h2 style= "color: #4b2102ff !important">☕ Đặt Bàn Online</h2>
        
        <!-- 2 NÚT CHỨC NĂNG -->
        <div class="btn-group mt-3">
            <a href="?url=reservation/create" class="btn btn-primary active">Đăng ký đặt bàn</a>
            <a href="?url=reservation/history" class="btn btn-outline-primary">Lịch sử đặt bàn</a>
        </div>
    </div>

    <div class="row justify-content-center">
        <div class="col-md-6">
            <div class="card shadow">
                <div class="card-body">
                    <form action="?url=reservation/store" method="POST">
                        <div class="mb-3">
                            <label>Họ và tên</label>
                            <input type="text" name="hoten" class="form-control" value="<?= $_SESSION['user']['name'] ?? '' ?>" required>
                        </div>
                        <div class="mb-3">
                            <label>Số điện thoại</label>
                            <input
                                type="text"
                                name="phone"
                                class="form-control"
                                value="<?= $_SESSION['user']['phone'] ?? '' ?>"
                                pattern="0[0-9]{9}"
                                maxlength="10"
                                title="Số điện thoại phải bắt đầu bằng 0 và có đúng 10 chữ số"
                                oninput="this.value = this.value.replace(/[^0-9]/g, '')"
                                required
                            >
                        </div>
                        <div class="row">
                            <div class="col-md-6 mb-3">
                                <label>Ngày đặt</label>
                                <input type="date" name="ngay" class="form-control" min="<?= date('Y-m-d') ?>" required>
                            </div>
                            <div class="col-md-6 mb-3">
                                <label>Giờ đến</label>
                                <input type="time" name="gio" class="form-control" required>
                            </div>
                        </div>
                        <div class="mb-3">
                            <label>Số lượng người</label>
                            <input type="number" name="songuoi" class="form-control" min="1" max="100" value="2" required>
                        </div>
                        <div class="mb-3">
                            <label>Ghi chú (Nếu có)</label>
                            <textarea name="ghichu" class="form-control" rows="3" placeholder="Ví dụ: Cần ghế trẻ em, ngồi gần cửa sổ..."></textarea>
                        </div>
                        <button type="submit" class="btn btn-success w-100">Gửi Yêu Cầu</button>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>

<?php require_once __DIR__ . '/../layouts/footer.php'; ?>  