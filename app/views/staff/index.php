<?php require_once __DIR__ . '/../layouts/admin_header.php'; ?>

<div class="container mt-4">
    <h2 class="mb-4 text-primary">☕ Quản lý Lịch Làm Việc Nhân Viên</h2>

    <div class="row">
        <!-- FORM XẾP LỊCH -->
        <div class="col-md-4">
            <div class="card shadow-sm">
                <div class="card-header bg-success text-white">
                    <i class="fas fa-plus-circle"></i> Xếp Lịch Mới
                </div>
                <div class="card-body">
                    <form action="?url=admin/staff/store" method="POST">
                        <div class="mb-3">
                            <label class="form-label">Chọn Nhân Viên</label>
                            <select name="user_id" class="form-select" required>
                                <?php foreach ($staffs as $staff): ?>
                                    <option value="<?= $staff['id'] ?>"><?= htmlspecialchars($staff['name']) ?> - ID: <?= $staff['id'] ?></option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Ngày làm</label>
                            <input type="date" name="work_date" class="form-control" value="<?= $selectedDate ?>" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Chọn Ca (25k/h)</label>
                            <select name="shift_code" class="form-select" required>
                                <option value="1">🌅 Ca Sáng (07:00 - 12:00)</option>
                                <option value="2">☀️ Ca Chiều (12:00 - 17:00)</option>
                                <option value="3">🌙 Ca Tối (17:00 - 22:00)</option>
                            </select>
                        </div>
                        <button type="submit" class="btn btn-primary w-100">Lưu Lịch Làm</button>
                    </form>
                </div>
            </div>
        </div>

        <!-- DANH SÁCH LỊCH LÀM THEO NGÀY -->
        <div class="col-md-8">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h4>Danh sách đi làm ngày: <span class="text-danger"><?= date('d/m/Y', strtotime($selectedDate)) ?></span></h4>
                
                <div class="d-flex gap-2">
                    <!-- Form chọn ngày -->
                    <form action="" method="GET" class="d-flex">
                        <input type="hidden" name="url" value="admin/staff">
                        <input type="date" name="date" class="form-control me-2" value="<?= $selectedDate ?>">
                        <button class="btn btn-secondary">Xem</button>
                    </form>

                    <!-- Nút Xuất Excel (Mới) -->
                    <a href="?url=admin/staff/export&date=<?= $selectedDate ?>" class="btn btn-success">
                        <i class="fas fa-file-excel"></i> Xuất Tuần
                    </a>
                </div>
            </div>

            <div class="card shadow-sm">
                <div class="card-body p-0">
                    <table class="table table-striped mb-0">
                        <thead class="table-dark">
                            <tr>
                                <th>Nhân viên</th>
                                <th>Ca làm</th>
                                <th>Thời gian</th>
                                <th>Lương dự kiến</th>
                                <th>Hành động</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php if (empty($dailyRoster)): ?>
                                <!-- ... Giữ nguyên ... -->
                            <?php else: ?>
                                <?php foreach ($dailyRoster as $row): 
                                    $hours = (strtotime($row['end_time']) - strtotime($row['start_time'])) / 3600;
                                    $salary = $hours * 25000;
                                ?>
                                <tr>
                                    <!-- ... (Cột Tên, Ca, Giờ, Lương giữ nguyên) ... -->
                                    <td>
                                        <div class="d-flex align-items-center">
                                            <img src="<?= !empty($row['avatar']) ? '/GocCaPhe/public/'.$row['avatar'] : '/GocCaPhe/public/assets/images/default-avatar.png' ?>" 
                                                 class="rounded-circle me-2" width="40" height="40" style="object-fit:cover;">
                                            <div>
                                                <strong><?= htmlspecialchars($row['name']) ?></strong><br>
                                                <small class="text-muted"><?= $row['phone'] ?></small>
                                            </div>
                                        </div>
                                    </td>
                                    <td>
                                        <?php if($row['shift_name']=='Ca Sáng') echo '<span class="badge bg-info">Sáng</span>'; ?>
                                        <?php if($row['shift_name']=='Ca Chiều') echo '<span class="badge bg-warning text-dark">Chiều</span>'; ?>
                                        <?php if($row['shift_name']=='Ca Tối') echo '<span class="badge bg-dark">Tối</span>'; ?>
                                    </td>
                                    <td><?= substr($row['start_time'],0,5) ?> - <?= substr($row['end_time'],0,5) ?></td>
                                    <td class="fw-bold text-success"><?= number_format($salary) ?> đ</td>
                                    
                                    <!-- CỘT HÀNH ĐỘNG ĐƯỢC CẬP NHẬT -->
                                    <td>
                                        <!-- Nút Sửa -->
                                        <a href="?url=admin/staff/edit&id=<?= $row['id'] ?>&date=<?= $selectedDate ?>" class="btn btn-sm btn-warning text-dark" title="Đổi người">
                                            <i class="fas fa-edit">Sửa</i>
                                        </a>
                                        <!-- Nút Xóa -->
                                        <a href="?url=admin/staff/delete&id=<?= $row['id'] ?>" class="btn btn-sm btn-danger" onclick="return confirm('Xóa lịch này?')">
                                            <i class="fas fa-trash">Xóa</i>
                                        </a>
                                    </td>
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