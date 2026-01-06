<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<div class="container mt-4 mb-5">
    <h2 class="text-primary mb-4"><i class="fas fa-chart-line"></i> Quản Lý Doanh Thu & Lợi Nhuận</h2>

    <div class="row">
        <!-- CỘT TRÁI: Dữ liệu chi tiết -->
        <div class="col-md-8">
            
            <!-- 1. BÁO CÁO NGÀY -->
            <div class="card shadow mb-4">
                <div class="card-header bg-info text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0"><i class="fas fa-calendar-day"></i> Doanh thu ngày: <?= date('d/m/Y', strtotime($filterDate)) ?></h5>
                    <form class="d-flex" action="index.php" method="GET">
                        <input type="hidden" name="url" value="admin/revenues">
                        <input type="hidden" name="month" value="<?= $filterMonth ?>">
                        <input type="hidden" name="year" value="<?= $filterYear ?>">
                        <input type="date" name="date" class="form-control form-control-sm me-2" value="<?= $filterDate ?>" onchange="this.form.submit()">
                    </form>
                </div>
                <div class="card-body">
                    <?php if ($dailyStats['total_orders'] == 0): ?>
                        <div class="text-center text-muted py-2">Chưa có đơn hàng trong ngày.</div>
                    <?php else: ?>
                        <div class="row text-center">
                            <div class="col-4">
                                <div class="p-2 border rounded bg-light">
                                    <small class="text-muted fw-bold">SỐ ĐƠN</small>
                                    <h4 class="mb-0 text-dark"><?= $dailyStats['total_orders'] ?></h4>
                                </div>
                            </div>
                            <div class="col-4">
                                <div class="p-2 border rounded bg-light">
                                    <small class="text-muted fw-bold">SỐ CỐC</small>
                                    <h4 class="mb-0 text-primary"><?= $dailyStats['total_cups'] ?></h4>
                                </div>
                            </div>
                            <div class="col-4">
                                <div class="p-2 border rounded bg-light border-success">
                                    <small class="text-success fw-bold">DOANH THU</small>
                                    <h4 class="mb-0 text-success"><?= number_format($dailyStats['total_revenue']) ?> ₫</h4>
                                </div>
                            </div>
                        </div>
                    <?php endif; ?>
                </div>
            </div>

            <!-- 2. BÁO CÁO THÁNG & CHI PHÍ (PHẦN MỚI) -->
            <div class="card shadow border-primary">
                <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
                    <h5 class="m-0"><i class="fas fa-calculator"></i> Tính Lãi/Lỗ Tháng <?= $filterMonth ?>/<?= $filterYear ?></h5>
                    <form class="d-flex" action="index.php" method="GET">
                        <input type="hidden" name="url" value="admin/revenues">
                        <select name="month" class="form-select form-select-sm me-1" onchange="this.form.submit()">
                            <?php for($m=1; $m<=12; $m++): ?>
                                <option value="<?= $m ?>" <?= $m == $filterMonth ? 'selected' : '' ?>>Tháng <?= $m ?></option>
                            <?php endfor; ?>
                        </select>
                        <select name="year" class="form-select form-select-sm" onchange="this.form.submit()">
                            <option value="2025" <?= $filterYear == 2025 ? 'selected' : '' ?>>2025</option>
                            <option value="2026" <?= $filterYear == 2026 ? 'selected' : '' ?>>2026</option>
                        </select>
                    </form>
                </div>
                <div class="card-body">
                    <div class="row">
                        <!-- FORM NHẬP CHI PHÍ -->
                        <div class="col-md-6 border-end">
                            <h6 class="text-danger fw-bold mb-3"><i class="fas fa-edit"></i> Nhập chi phí tháng:</h6>
                            <form action="index.php?url=admin/revenues&month=<?= $filterMonth ?>&year=<?= $filterYear ?>" method="POST">
                                <input type="hidden" name="action" value="save_expenses">
                                
                                <div class="mb-2">
                                    <label class="small text-muted fw-bold">1. Tiền Nhân viên:</label>
                                    <div class="input-group input-group-sm">
                                        <input type="number" name="staff_cost" class="form-control" placeholder="0" value="<?= $expenses['staff_cost'] ?>">
                                        <span class="input-group-text">₫</span>
                                    </div>
                                </div>

                                <div class="mb-2">
                                    <label class="small text-muted fw-bold">2. Tiền Điện / Nước:</label>
                                    <div class="input-group input-group-sm">
                                        <input type="number" name="electric_cost" class="form-control" placeholder="0" value="<?= $expenses['electricity_water_cost'] ?>">
                                        <span class="input-group-text">₫</span>
                                    </div>
                                </div>

                                <div class="mb-3">
                                    <label class="small text-muted fw-bold">3. Tiền Nguyên Vật Liệu:</label>
                                    <div class="input-group input-group-sm">
                                        <input type="number" name="material_cost" class="form-control" placeholder="0" value="<?= $expenses['materials_cost'] ?>">
                                        <span class="input-group-text">₫</span>
                                    </div>
                                </div>

                                <button type="submit" class="btn btn-outline-primary btn-sm w-100">
                                    <i class="fas fa-save"></i> Cập nhật chi phí
                                </button>
                            </form>
                        </div>

                        <!-- KẾT QUẢ TÍNH TOÁN -->
                        <div class="col-md-6">
                            <h6 class="text-success fw-bold mb-3"><i class="fas fa-money-bill-wave"></i> Kết quả kinh doanh:</h6>
                            <ul class="list-group list-group-flush mb-3">
                                <li class="list-group-item d-flex justify-content-between align-items-center">
                                    <span>Tổng Doanh Thu:</span>
                                    <span class="fw-bold text-success">+ <?= number_format($monthlyRevenue) ?> ₫</span>
                                </li>
                                <li class="list-group-item d-flex justify-content-between align-items-center bg-light">
                                    <span>Tổng Chi Phí:</span>
                                    <span class="fw-bold text-danger">- <?= number_format($totalExpenses) ?> ₫</span>
                                </li>
                            </ul>

                            <div class="alert <?= $realProfit >= 0 ? 'alert-success' : 'alert-danger' ?> text-center">
                                <small class="text-uppercase fw-bold">Lợi nhuận thực tế</small>
                                <h2 class="fw-bold my-1">
                                    <?= number_format($realProfit) ?> ₫
                                </h2>
                                <small><?= $realProfit >= 0 ? '(Đang có lãi)' : '(Đang lỗ)' ?></small>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="card-footer bg-light text-center">
                    <a href="index.php?url=admin/revenues/export&month=<?= $filterMonth ?>&year=<?= $filterYear ?>" class="btn btn-success fw-bold">
                        <i class="fas fa-file-excel"></i> Xuất Báo Cáo Lợi Nhuận
                    </a>
                </div>
            </div>
        </div>

        <!-- CỘT PHẢI: BIỂU ĐỒ -->
        <div class="col-md-4">
            <div class="card shadow h-100">
                <div class="card-header bg-white font-weight-bold">
                    <i class="fas fa-chart-pie text-warning"></i> Tỷ trọng doanh thu
                </div>
                <div class="card-body d-flex flex-column justify-content-center">
                    <?php if (empty($chartValues) || array_sum($chartValues) == 0): ?>
                        <div class="text-center text-muted">
                            <i class="fas fa-chart-pie fa-3x mb-3" style="color: #ddd;"></i>
                            <p>Chưa có số liệu.</p>
                        </div>
                    <?php else: ?>
                        <canvas id="revenueChart"></canvas>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>
</div>

<?php if (!empty($chartValues)): ?>
<script>
    const ctx = document.getElementById('revenueChart').getContext('2d');
    new Chart(ctx, {
        type: 'doughnut', 
        data: {
            labels: <?= json_encode($chartLabels) ?>,
            datasets: [{
                data: <?= json_encode($chartValues) ?>,
                backgroundColor: ['#4e73df', '#1cc88a', '#36b9cc', '#f6c23e', '#e74a3b', '#858796', '#20c9a6', '#6610f2'],
                hoverOffset: 4
            }]
        },
        options: {
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.label + ': ' + new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(context.raw);
                        }
                    }
                }
            }
        }
    });
</script>
<?php endif; ?>