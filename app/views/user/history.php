<?php require_once __DIR__ . '/../layouts/header.php'; ?>

<div class="container mt-5 mb-5">
    <div class="text-center mb-4">
        <h2 class="text-primary">📅 Lịch Sử Đặt Bàn Của Tôi</h2>

        <!-- 2 NÚT CHỨC NĂNG -->
        <div class="btn-group mt-3">
            <a href="?url=reservation/create" class="btn btn-outline-primary">Đăng ký đặt bàn</a>
            <a href="?url=reservation/history" class="btn btn-primary active">Lịch sử đặt bàn</a>
        </div>
    </div>

    <div class="card shadow-sm">
        <div class="card-body">
            <?php if (empty($myReservations)): ?>
                <div class="text-center py-4">
                    <p class="text-muted">Bạn chưa có lịch đặt bàn nào.</p>
                    <a href="?url=reservation/create" class="btn btn-sm btn-success">Đặt ngay</a>
                </div>
            <?php else: ?>
                <table class="table table-striped table-hover align-middle">
                    <thead class="table-dark">
                        <tr>
                            <th>Ngày & Giờ</th>
                            <th>Số người</th>
                            <th>Thông tin liên hệ</th>
                            <th>Ghi chú</th>
                            <th>Trạng thái</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($myReservations as $r): ?>
                        <tr>
                            <td>
                                <div class="fw-bold"><?= date('d/m/Y', strtotime($r['ngay'])) ?></div>
                                <div class="text-primary"><?= substr($r['gio'], 0, 5) ?></div>
                            </td>
                            <td><?= $r['songuoi'] ?> khách</td>
                            <td>
                                <div><?= htmlspecialchars($r['hoten']) ?></div>
                                <small class="text-muted"><?= $r['phone'] ?></small>
                            </td>
                            <td><?= htmlspecialchars($r['ghichu']) ?></td>
                            <td>
                                <?php 
                                    if ($r['trangthai'] == 'DA_DUYET') echo '<span class="badge bg-success">Thành công</span>';
                                    elseif ($r['trangthai'] == 'HUY') echo '<span class="badge bg-danger">Đã hủy</span>';
                                    else echo '<span class="badge bg-warning text-dark">Chờ duyệt</span>';
                                ?>
                            </td>
                        </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            <?php endif; ?>
        </div>
    </div>
</div>

<?php require_once __DIR__ . '/../layouts/footer.php'; ?>