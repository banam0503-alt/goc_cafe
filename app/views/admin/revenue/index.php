<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>
<!-- Nhúng Chart.js -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<div class="container mt-4 mb-5">
    <h2 class="text-primary mb-4"><i class="fas fa-chart-line"></i> Quản Lý Doanh Thu</h2>

    <div class="row">
        <!-- CỘT TRÁI: SỐ LIỆU -->
        <div class="col-md-8">
            
            <!-- 1. BÁO CÁO NGÀY -->
            <div class="card shadow mb-4">
                <div class="card-header bg-info text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0">
                        <i class="fas fa-calendar-day"></i> 
                        Doanh thu ngày: <?= date('d/m/Y', strtotime($filterDate)) ?>
                    </h5>
                    
                    <form class="d-flex" action="index.php" method="GET">
                        <input type="hidden" name="url" value="admin/revenues">
                        <input type="hidden" name="month" value="<?= $filterMonth ?>">
                        <input type="hidden" name="year" value="<?= $filterYear ?>">
                        
                        <input type="date" name="date" class="form-control form-control-sm me-2" value="<?= $filterDate ?>" onchange="this.form.submit()">
                        <button type="submit" class="btn btn-light btn-sm fw-bold">Xem</button>
                    </form>
                </div>
                <div class="card-body">
                    <?php if ($dailyStats['total_orders'] == 0): ?>
                        <div class="text-center py-3 text-muted">
                            <i class="fas fa-mug-hot fa-3x mb-2"></i><br>
                            Chưa có đơn hàng nào hoàn thành trong ngày này.
                        </div>
                    <?php else: ?>
                        <div class="row text-center">
                            <div class="col-md-4 mb-3">
                                <div class="p-3 border rounded bg-light shadow-sm">
                                    <small class="text-muted text-uppercase fw-bold">Đơn hàng</small>
                                    <h4 class="fw-bold text-dark mt-2"><?= $dailyStats['total_orders'] ?></h4>
                                </div>
                            </div>
                            <div class="col-md-4 mb-3">
                                <div class="p-3 border rounded bg-light shadow-sm">
                                    <small class="text-muted text-uppercase fw-bold">Số cốc bán</small>
                                    <h4 class="fw-bold text-primary mt-2"><?= $dailyStats['total_cups'] ?></h4>
                                </div>
                            </div>
                            <div class="col-md-4 mb-3">
                                <div class="p-3 border rounded bg-light shadow-sm border-success">
                                    <small class="text-success text-uppercase fw-bold">Tổng Doanh thu</small>
                                    <h4 class="fw-bold text-success mt-2"><?= number_format($dailyStats['total_revenue']) ?> ₫</h4>
                                </div>
                            </div>
                        </div>
                    <?php endif; ?>
                </div>
            </div>

            <!-- 2. BÁO CÁO THÁNG (Đã sửa lỗi hiển thị sai biến) -->
            <div class="card shadow border-primary">
                <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0"><i class="fas fa-calendar-alt"></i> Tháng <?= $filterMonth ?>/<?= $filterYear ?></h5>
                    
                     <div class="d-flex gap-1">
                        <form class="d-flex" action="index.php" method="GET">
                            <input type="hidden" name="url" value="admin/revenues">
                            <input type="hidden" name="date" value="<?= $filterDate ?>">
                            
                            <select name="month" class="form-select form-select-sm me-2" style="width: 70px;" onchange="this.form.submit()">
                                <?php for($m=1; $m<=12; $m++): ?>
                                    <option value="<?= $m ?>" <?= $m == $filterMonth ? 'selected' : '' ?>>T<?= $m ?></option>
                                <?php endfor; ?>
                            </select>
                            <select name="year" class="form-select form-select-sm me-2" style="width: 80px;" onchange="this.form.submit()">
                                <option value="2024" <?= $filterYear == 2024 ? 'selected' : '' ?>>2024</option>
                                <option value="2025" <?= $filterYear == 2025 ? 'selected' : '' ?>>2025</option>
                                <option value="2026" <?= $filterYear == 2026 ? 'selected' : '' ?>>2026</option>
                            </select>
                            <button class="btn btn-light btn-sm fw-bold">Xem</button>
                        </form>
                            <!-- NÚT XUẤT EXCEL (MỚI) -->
                            <a href="index.php?url=admin/revenues/export&month=<?= $filterMonth ?>&year=<?= $filterYear ?>" 
                            class="btn btn-warning btn-sm fw-bold" title="Xuất Excel">
                                <i class="fas fa-file-excel"></i> Xuất
                            </a>
                    </div>
                </div>
                <div class="card-body text-center">
                    <h6 class="text-muted text-uppercase mb-3">Tổng doanh thu thực tế (Đã trừ đơn hủy/chưa trả)</h6>
                    <!-- SỬA Ở ĐÂY: Dùng biến $monthlyStats['revenue'] -->
                    <h1 class="display-4 fw-bold text-success">
                        <?= number_format($monthlyStats['revenue']) ?> ₫
                    </h1>
                </div>
            </div>
        </div>

        <!-- CỘT PHẢI: BIỂU ĐỒ (Giữ nguyên) -->
        <div class="col-md-4">
            <div class="card shadow h-100">
                <div class="card-header bg-white font-weight-bold">
                    <i class="fas fa-chart-pie text-warning"></i> Tỷ trọng doanh thu (T<?= $filterMonth ?>)
                </div>
                <div class="card-body d-flex flex-column justify-content-center">
                    <?php if (empty($chartValues) || array_sum($chartValues) == 0): ?>
                        <div class="text-center text-muted">
                            <i class="fas fa-chart-pie fa-3x mb-3" style="color: #ddd;"></i>
                            <p>Chưa có đơn hàng nào trong tháng này.</p>
                        </div>
                    <?php else: ?>
                        <canvas id="revenueChart"></canvas>
                        <div class="mt-4 small text-muted text-center">
                            * Biểu đồ dựa trên tổng tiền bán được của từng nhóm món ăn.
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- SCRIPT (Giữ nguyên) -->
<?php if (!empty($chartValues)): ?>
<script>
    const ctx = document.getElementById('revenueChart').getContext('2d');
    new Chart(ctx, {
        type: 'doughnut', 
        data: {
            labels: <?= json_encode($chartLabels) ?>,
            datasets: [{
                data: <?= json_encode($chartValues) ?>,
                backgroundColor: [
                    '#4e73df', '#1cc88a', '#36b9cc', '#f6c23e', '#e74a3b', '#858796', '#20c9a6', '#6610f2'
                ],
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let val = new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(context.raw);
                            return context.label + ': ' + val;
                        }
                    }
                }
            }
        }
    });
</script>
<?php endif; ?>