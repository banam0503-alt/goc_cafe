<!-- views/staff/edit.php -->
<?php require_once __DIR__ . '/../layouts/admin_header.php'; 
// Xác định shift_code hiện tại để auto selected
$currentCode = 1;
if ($shift['shift_name'] == 'Ca Chiều') $currentCode = 2;
if ($shift['shift_name'] == 'Ca Tối') $currentCode = 3;
?>

<div class="container mt-5">
    <div class="row justify-content-center">
        <div class="col-md-6">
            <div class="card shadow">
                <div class="card-header bg-warning text-dark fw-bold">
                    <i class="fas fa-edit"></i> Chỉnh Sửa Lịch Làm Việc
                </div>
                <div class="card-body">
                    <form action="?url=admin/staff/update" method="POST">
                        <input type="hidden" name="schedule_id" value="<?= $shift['id'] ?>">
                        <input type="hidden" name="current_date" value="<?= $_GET['date'] ?? $shift['work_date'] ?>">

                        <!-- 1. Chọn Nhân Viên -->
                        <div class="mb-3">
                            <label class="form-label fw-bold">Nhân viên:</label>
                            <select name="user_id" class="form-select" required>
                                <?php foreach ($staffs as $staff): ?>
                                    <option value="<?= $staff['id'] ?>" <?= ($staff['id'] == $shift['user_id']) ? 'selected' : '' ?>>
                                        <?= htmlspecialchars($staff['name']) ?> (ID: <?= $staff['id'] ?>)
                                    </option>
                                <?php endforeach; ?>
                            </select>
                        </div>

                        <!-- 2. Chọn Ca Làm (Mới thêm) -->
                        <div class="mb-3">
                            <label class="form-label fw-bold">Ca làm việc:</label>
                            <select name="shift_code" class="form-select" required>
                                <option value="1" <?= $currentCode == 1 ? 'selected' : '' ?>>🌅 Ca Sáng (07:00 - 12:00)</option>
                                <option value="2" <?= $currentCode == 2 ? 'selected' : '' ?>>☀️ Ca Chiều (12:00 - 17:00)</option>
                                <option value="3" <?= $currentCode == 3 ? 'selected' : '' ?>>🌙 Ca Tối (17:00 - 22:00)</option>
                            </select>
                        </div>

                        <div class="d-flex justify-content-between mt-4">
                            <a href="?url=admin/staff&date=<?= $_GET['date'] ?? $shift['work_date'] ?>" class="btn btn-secondary">Quay lại</a>
                            <button type="submit" class="btn btn-primary">Lưu thay đổi</button>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>