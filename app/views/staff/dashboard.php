<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Staff Dashboard</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .staff-header { background: #7a4a2e; color: white; padding: 15px 0; margin-bottom: 30px; }
        .feature-box { text-align: center; padding: 20px; border: 1px solid #eee; border-radius: 10px; transition: 0.3s; background: #fff; cursor: pointer; }
        .feature-box:hover { transform: translateY(-5px); box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .feature-box i { font-size: 40px; margin-bottom: 15px; color: #28a745; }
        .feature-box h5 { font-weight: bold; color: #333; }
        .schedule-table th { background-color: #f8f9fa; }
    </style>
</head>
<body style="background-color: #f4f6f9;">

<!-- HEADER STAFF -->
<div class="staff-header shadow">
    <div class="container d-flex justify-content-between align-items-center">

        <h3 class="m-0"><i ></i> Xin chào, <?= htmlspecialchars($_SESSION['user']['name']) ?></h3>
        <div>
            <a href="?url=logout" class="btn btn-light btn-sm fw-bold">Đăng xuất</a>
        </div>
        
    </div>
</div>

<div class="container">
    <div class="row">
        <!-- CỘT TRÁI: MENU CHỨC NĂNG -->
        <div class="col-md-4 mb-4">
            <h5 class="mb-3 text-uppercase text-muted">Chức năng</h5>
            
            <div class="row g-3">
                <!-- 1. Cập nhật Profile -->
                <div class="">
                    <a href="?url=profile" class="text-decoration-none">
                        <div class="feature-box">
                            <i class="fas fa-id-card"></i>
                            <h5>Hồ sơ</h5>
                            <small>Xem & Sửa</small>
                        </div>
                    </a>
                </div>
                <!-- 2. Duyệt Đơn -->
                <div class="">
                    <a href="?url=admin/orders" class="text-decoration-none">
                        <div class="feature-box">
                            <i class="fas fa-receipt"></i>
                            <h5>Duyệt Đơn</h5>
                            <small>Quản lý đơn</small>
                        </div>
                    </a>
                </div>
                <!-- 3. Duyệt Bàn -->
                <div class="">
                    <a href="?url=admin/reservations" class="text-decoration-none">
                        <div class="feature-box">
                            <i class="fas fa-chair"></i>
                            <h5>Duyệt Bàn</h5>
                            <small>Đặt chỗ</small>
                        </div>
                    </a>
                </div>
                <!-- 4. Xem trang chủ -->
                <div class="">
                    <a href="/GocCaPhe/public/index.php" class="text-decoration-none">
                        <div class="feature-box">
                            <i class="fas fa-home"></i>
                            <h5>Trang chủ</h5>
                            <small>Về trang web</small>
                        </div>
                    </a>
                </div>
            </div>
        </div>

        <!-- CỘT PHẢI: LỊCH LÀM VIỆC CÁ NHÂN -->
        <div class="col-md-8">
            <div class="card border-0 shadow-sm">
                <div class="card-header bg-white border-bottom py-3">
                    <h5 class="m-0 text-success"><i class="fas fa-calendar-alt"></i> Lịch Làm Việc Của Tôi</h5>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table schedule-table mb-0">
                            <thead>
                                <tr>
                                    <th>Ngày</th>
                                    <th>Ca làm</th>
                                    <th>Thời gian</th>
                                    <th>Lương (Dự kiến)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <?php if (empty($mySchedules)): ?>
                                    <tr><td colspan="4" class="text-center py-4 text-muted">Bạn chưa có lịch làm việc sắp tới.</td></tr>
                                <?php else: ?>
                                    <?php foreach ($mySchedules as $sch): 
                                         $salary = 5 * 25000; // 5 tiếng * 25k
                                         $date = date('d/m/Y', strtotime($sch['work_date']));
                                         $today = date('d/m/Y');
                                    ?>
                                    <tr class="<?= ($date == $today) ? 'table-warning' : '' ?>">
                                        <td class="fw-bold">
                                            <?= $date ?>
                                            <?= ($date == $today) ? '<span class="badge bg-danger">Hôm nay</span>' : '' ?>
                                        </td>
                                        <td><?= $sch['shift_name'] ?></td>
                                        <td><?= substr($sch['start_time'], 0, 5) ?> - <?= substr($sch['end_time'], 0, 5) ?></td>
                                        <td class="text-success fw-bold"><?= number_format($salary) ?> đ</td>
                                    </tr>
                                    <?php endforeach; ?>
                                <?php endif; ?>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>