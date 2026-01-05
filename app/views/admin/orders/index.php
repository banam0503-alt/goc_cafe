<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>

<div class="container mt-4 mb-5">
    <h3 class="text-primary"><i class="fas fa-file-invoice-dollar"></i> Quản lý đơn hàng</h3>

    <div class="d-flex justify-content-between align-items-center mb-3">
    <div></div>

    <a href="?url=admin/orders/export-excel"
       class="btn btn-success"
       title="Xuất Excel đơn hàng đã duyệt">
        <i class="fas fa-file-excel"></i> Xuất Excel
    </a>
</div>

    <div class="card shadow-sm">
        <div class="card-body p-0">
            <table class="table table-bordered table-hover align-middle mb-0">
                <thead class="table-dark">
                    <tr>
                        <th>ID</th>
                        <th>Khách hàng</th>
                        <th>Tổng tiền</th>
                        <th>Trạng thái</th>
                        <th>Người xử lý</th>
                        <th>Ngày tạo</th>
                        <th width="150">Hành động</th>
                    </tr>
                </thead>
                <tbody>
                <?php foreach($orders as $o): ?>
                    <tr>
                        <td><strong>#<?= $o['id'] ?></strong></td>
                        
                        <!-- Thông tin khách hàng -->
                        <td>
                            <?= htmlspecialchars($o['customer_name']) ?>
                        </td>

                        <!-- Tổng tiền -->
                        <td class="fw-bold text-success">
                            <?= number_format($o['total_price'],0,',','.') ?>₫
                        </td>

                        <!-- Trạng thái -->
                        <td>
                            <?php 
                                $statusMap = [
                                    'PAID'      => ['bg' => 'warning text-dark', 'label' => 'Chờ duyệt'],
                                    'APPROVED'  => ['bg' => 'success', 'label' => 'Đã duyệt'],
                                    'CANCELLED' => ['bg' => 'danger', 'label' => 'Đã hủy'],
                                    'SHIPPING'  => ['bg' => 'info text-dark', 'label' => 'Đang giao'],
                                    'COMPLETED' => ['bg' => 'primary', 'label' => 'Hoàn thành']
                                ];
                                $st = $statusMap[$o['status']] ?? ['bg'=>'secondary', 'label'=>$o['status']];
                            ?>
                            <span class="badge bg-<?= $st['bg'] ?>">
                                <?= $st['label'] ?>
                            </span>
                            
                            <!-- Nếu bị hủy thì hiện lý do -->
                            <?php if($o['status'] === 'CANCELLED' && !empty($o['reject_reason'])): ?>
                                <div class="small text-danger mt-1">
                                    <i class="fas fa-comment-alt"></i> Lý do: <?= htmlspecialchars($o['reject_reason']) ?>
                                </div>
                            <?php endif; ?>
                        </td>

                        <!-- Người xử lý (Hiển thị Tên + Vai trò) -->
                        <!-- Người xử lý (Hiển thị Tên + Vai trò) -->
                        <td>
                            <?php if(!empty($o['staff_name'])): ?>
                                <div><i class="fas fa-user-check"></i> <?= htmlspecialchars($o['staff_name']) ?></div>
                                
                                <?php 
                                    // Logic chọn màu dựa trên Role
                                    $roleClass = match($o['staff_role']) {
                                        'ADMIN' => 'bg-success',           // Màu Xanh lá
                                        'STAFF' => 'bg-warning text-dark', // Màu Vàng (thêm text-dark để chữ đen dễ đọc)
                                        default => 'bg-secondary'          // Mặc định màu Xám
                                    };
                                ?>
                                
                                <span class="badge <?= $roleClass ?>" style="font-size: 0.7rem;">
                                    <?= $o['staff_role'] ?>
                                </span>

                                <div class="small text-muted">
                                    <?= date('H:i d/m', strtotime($o['approved_at'])) ?>
                                </div>
                            <?php else: ?>
                                <span class="text-muted">—</span>
                            <?php endif; ?>
                        </td>

                        <!-- Thời gian tạo -->
                        <td><?= date('d/m/Y H:i', strtotime($o['created_at'])) ?></td>

                        <!-- Hành động -->
                        <td>
                            <?php if($o['status'] == 'PAID'): ?>
                                <div class="d-flex gap-1">
                                    <form method="post" action="?url=admin/orders/approve">
                                        <input type="hidden" name="id" value="<?= $o['id'] ?>">
                                        <button class="btn btn-sm btn-success" title="Duyệt đơn">
                                            <i class="fas fa-check"></i> Duyệt
                                        </button>
                                    </form>

                                    <button class="btn btn-sm btn-danger" 
                                            onclick="rejectOrder(<?= $o['id'] ?>)" 
                                            title="Từ chối">
                                        <i class="fas fa-times"></i> Hủy
                                    </button>
                                </div>
                            <?php else: ?>
                                <span class="text-muted small"><i class="fas fa-lock"></i> Đã đóng</span>
                            <?php endif; ?>
                        </td>
                    </tr>
                <?php endforeach; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>  

<script>
function rejectOrder(id){
    const reason = prompt('Vui lòng nhập lý do từ chối đơn hàng này:');
    
    // Nếu người dùng ấn Cancel hoặc không nhập gì
    if(reason === null) return; 
    if(reason.trim() === "") {
        alert("Bạn phải nhập lý do!");
        return;
    }

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '?url=admin/orders/reject';

    form.innerHTML = `
        <input type="hidden" name="id" value="${id}">
        <input type="hidden" name="reject_reason" value="${reason}"> 
    `;
    // Lưu ý: name="reject_reason" phải khớp với controller

    document.body.appendChild(form);
    form.submit();
}
</script>