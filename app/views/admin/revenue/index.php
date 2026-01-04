<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>
<!-- Nhúng Chart.js -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<div class="container mt-4 mb-5">
    <h2 class="text-primary mb-4"><i class="fas fa-chart-line"></i> Quản Lý Doanh Thu</h2>

    <div class="row">
        <!-- CỘT TRÁI: SỐ LIỆU -->
        <div class="col-md-8">
            
            <!-- 1. BÁO CÁO NGÀY (FORM SỬA LỖI TẠI ĐÂY) -->
            <div class="card shadow mb-4">
                <div class="card-header bg-info text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0">
                        <i class="fas fa-calendar-day"></i> 
                        Ngày: <?= date('d/m/Y', strtotime($filterDate)) ?>
                    </h5>
                    
                    <!-- QUAN TRỌNG: Action để trống hoặc trỏ về index.php -->
                    <form class="d-flex" action="index.php" method="GET">
                        <!-- 1. Bắt buộc phải có dòng này để giữ đường dẫn admin/revenues -->
                        <input type="hidden" name="url" value="admin/revenues">
                        
                        <!-- 2. Giữ lại tháng/năm hiện tại để không làm hỏng biểu đồ bên dưới -->
                        <input type="hidden" name="month" value="<?= $filterMonth ?>">
                        <input type="hidden" name="year" value="<?= $filterYear ?>">
                        
                        <!-- 3. Value phải in ra biến $filterDate để sau khi reload nó vẫn hiện ngày đã chọn -->
                        <input type="date" name="date" class="form-control form-control-sm me-2" value="<?= $filterDate ?>" onchange="this.form.submit()">
                        
                        <button type="submit" class="btn btn-light btn-sm fw-bold">Xem</button>
                    </form>
                </div>
                <div class="card-body">
                    <?php if ($dailyStats['total_orders'] == 0): ?>
                        <div class="text-center py-3 text-muted">
                            <i class="fas fa-mug-hot fa-3x mb-2"></i><br>
                            Không có đơn hàng nào trong ngày này.
                        </div>
                    <?php else: ?>
                        <div class="row text-center">
                            <div class="col-md-4 mb-3">
                                <div class="p-2 border rounded bg-light">
                                    <small class="text-muted">Đơn hàng</small>
                                    <h4 class="fw-bold text-dark"><?= $dailyStats['total_orders'] ?></h4>
                                </div>
                            </div>
                            <div class="col-md-4 mb-3">
                                <div class="p-2 border rounded bg-light">
                                    <small class="text-muted">Số cốc bán</small>
                                    <h4 class="fw-bold text-primary"><?= $dailyStats['total_cups'] ?></h4>
                                </div>
                            </div>
                            <div class="col-md-4 mb-3">
                                <div class="p-2 border rounded bg-light">
                                    <small class="text-muted">Doanh thu</small>
                                    <h4 class="fw-bold text-success"><?= number_format($dailyStats['total_revenue']) ?>đ</h4>
                                </div>
                            </div>
                        </div>
                        <div class="alert alert-secondary p-2 m-0 text-center">
                            Lợi nhuận (Doanh thu - Lương NV): 
                            <strong class="<?= $dailyProfit >= 0 ? 'text-success' : 'text-danger' ?>">
                                <?= number_format($dailyProfit) ?> đ
                            </strong>
                            <div style="font-size: 12px;" class="text-muted mt-1">
                                (Lương hôm nay: <?= number_format($dailySalary) ?>đ)
                            </div>
                        </div>
                    <?php endif; ?>
                </div>
            </div>

            <!-- 2. BÁO CÁO THÁNG -->
            <div class="card shadow border-primary">
                <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0"><i class="fas fa-calendar-alt"></i> Tháng <?= $filterMonth ?>/<?= $filterYear ?></h5>
                    
                    <form class="d-flex" action="index.php" method="GET">
                        <input type="hidden" name="url" value="admin/revenues">
                        <!-- Giữ lại ngày đang chọn ở trên -->
                        <input type="hidden" name="date" value="<?= $filterDate ?>">
                        
                        <select name="month" class="form-select form-select-sm me-2" style="width: 70px;" onchange="this.form.submit()">
                            <?php for($m=1; $m<=12; $m++): ?>
                                <option value="<?= $m ?>" <?= $m == $filterMonth ? 'selected' : '' ?>>T<?= $m ?></option>
                            <?php endfor; ?>
                        </select>
                        <select name="year" class="form-select form-select-sm me-2" style="width: 80px;" onchange="this.form.submit()">
                            <option value="2025" <?= $filterYear == 2025 ? 'selected' : '' ?>>2025</option>
                            <option value="2026" <?= $filterYear == 2026 ? 'selected' : '' ?>>2026</option>
                        </select>
                        <button class="btn btn-light btn-sm fw-bold">Xem</button>
                    </form>
                </div>
                <div class="card-body">
                    <table class="table table-bordered mb-0">
                        <tr>
                            <td>Tổng Doanh Thu:</td>
                            <td class="text-end fw-bold text-success"><?= number_format($monthlyStats['revenue']) ?> đ</td>
                        </tr>
                        <tr>
                            <td>Tổng Lương NV:</td>
                            <td class="text-end fw-bold text-danger">- <?= number_format($monthlyStats['salary']) ?> đ</td>
                        </tr>
                        <tr class="table-active">
                            <td><strong>Lợi Nhuận Thực:</strong></td>
                            <td class="text-end fw-bold fs-5 <?= $monthlyProfit >= 0 ? 'text-success' : 'text-danger' ?>">
                                <?= number_format($monthlyProfit) ?> đ
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>

        <!-- CỘT PHẢI: BIỂU ĐỒ TRÒN -->
        <div class="col-md-4">
            <div class="card shadow h-100">
                <div class="card-header bg-white font-weight-bold">
                    <i class="fas fa-chart-pie text-warning"></i> Doanh thu theo món (T<?= $filterMonth ?>)
                </div>
                <div class="card-body d-flex flex-column justify-content-center">
                    <?php if (empty($chartValues) || array_sum($chartValues) == 0): ?>
                        <div class="text-center text-muted">
                            <i class="fas fa-chart-pie fa-3x mb-3" style="color: #ddd;"></i>
                            <p>Chưa có dữ liệu bán hàng tháng này.</p>
                        </div>
                    <?php else: ?>
                        <canvas id="revenueChart"></canvas>
                        <div class="mt-4 small text-muted text-center">
                            * Biểu đồ thể hiện tỷ trọng doanh thu giữa các loại đồ uống trong tháng.
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- SCRIPT VẼ BIỂU ĐỒ -->
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
                    '#4e73df', '#1cc88a', '#36b9cc', '#f6c23e', '#e74a3b', '#858796', '#20c9a6'
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