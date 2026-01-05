<?php require_once __DIR__ . '/../layouts/header.php'; ?>
<style>
/* Xóa nút tăng giảm mặc định trên Chrome, Safari, Edge và Opera */
input::-webkit-outer-spin-button,
input::-webkit-inner-spin-button {
    -webkit-appearance: none;
    margin: 0;
}

/* Xóa nút tăng giảm mặc định trên Firefox */
input[type=number] {
    -moz-appearance: textfield;
}
</style>
<div class="container mt-4">
    <h2>Giỏ hàng của bạn</h2>

    <div class="mb-3">
        <a href="?url=cart&status=PENDING"
           class="btn btn-outline-warning <?= ($_GET['status'] ?? 'PENDING') === 'PENDING' ? 'active' : '' ?>">
            Đang mua (PENDING)
        </a>
        <a href="?url=cart&status=PAID"
           class="btn btn-outline-info <?= ($_GET['status'] ?? '') === 'PAID' ? 'active' : '' ?>">
            Đã thanh toán (PAID)
        </a>
        <a href="?url=cart&status=APPROVED"
           class="btn btn-outline-success <?= ($_GET['status'] ?? '') === 'APPROVED' ? 'active' : '' ?>">
            Đã duyệt
        </a>
        <a href="?url=cart&status=CANCELLED"
           class="btn btn-outline-danger <?= ($_GET['status'] ?? '') === 'CANCELLED' ? 'active' : '' ?>">
            Đã hủy
        </a>
    </div>

    <?php if (empty($items)): ?>
        <div class="alert alert-light border">Giỏ hàng trong mục này đang trống</div>
    <?php else: ?>
        <form id="cart-form">
            <table class="table table-striped table-bordered align-middle">
                <thead class="table-dark">
                    <tr>
                        <th><input type="checkbox" id="select-all"></th>
                        <th>Tên sản phẩm</th>
                        <th>Số lượng</th>
                        <th>Giá</th>
                        <th>Trạng thái</th>
                        <th>Tạm tính</th>
                        <th>Ảnh</th>
                        <th class="text-center">Action</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($items as $item): ?>
                        <tr data-order-item-id="<?= $item['order_item_id'] ?>">
                            <td>
                                <input type="checkbox" class="item-checkbox" value="<?= $item['order_item_id'] ?>" 
                                <?= $item['order_status'] !== 'PENDING' ? 'disabled' : '' ?>>
                            </td>
                            <td><?= htmlspecialchars($item['name']) ?></td>
                            <td>
                                <div class="input-group" style="width:120px;">
                                    <button type="button" class="btn btn-outline-secondary btn-decrease" <?= $item['order_status'] !== 'PENDING' ? 'disabled' : '' ?>>-</button>
                                    <input type="number" class="form-control text-center quantity-input" value="<?= $item['quantity'] ?>" min="1" max="50" <?= $item['order_status'] !== 'PENDING' ? 'readonly' : '' ?>>
                                    <button type="button" class="btn btn-outline-secondary btn-increase" <?= $item['order_status'] !== 'PENDING' ? 'disabled' : '' ?>>+</button>
                                </div>
                            </td>
                            <td class="price"><?= number_format($item['price'], 0, ',', '.') ?>₫</td>
                            <td>
                                <?php
                                $badgeClass = match ($item['order_status']) {
                                    'PENDING'   => 'bg-warning',
                                    'PAID'      => 'bg-info',
                                    'APPROVED'  => 'bg-success',
                                    'CANCELLED' => 'bg-danger',
                                    default     => 'bg-secondary'
                                };
                                ?>
                                <span class="badge <?= $badgeClass ?>"><?= $item['order_status'] ?></span>
                            </td>
                            <td class="subtotal"><?= number_format($item['price'] * $item['quantity'], 0, ',', '.') ?>₫</td>
                            <td class="text-center">
                                <?php if (!empty($item['image'])): ?>
                                    <button type="button" class="btn btn-sm btn-info" data-bs-toggle="modal" data-bs-target="#imgModal<?= $item['order_item_id'] ?>">
                                        👁️
                                    </button>

                                    <div class="modal fade" id="imgModal<?= $item['order_item_id'] ?>" tabindex="-1" aria-hidden="true">
                                        <div class="modal-dialog modal-dialog-centered">
                                            <div class="modal-content">
                                                <div class="modal-header">
                                                    <h5 class="modal-title"><?= htmlspecialchars($item['name']) ?></h5>
                                                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                                                </div>
                                                <div class="modal-body text-center">
                                                    <?php 
                                                        // Đường dẫn chuẩn dựa theo ảnh cấu trúc thư mục của bạn
                                                        $imagePath = "/GocCaPhe/public/assets/img/" . $item['image']; 
                                                    ?>
                                                    <img src="<?= $imagePath ?>" 
                                                        class="img-fluid rounded" 
                                                        alt="Sản phẩm"
                                                        onerror="this.src='/GocCaPhe/public/assets/img/default.jpg';"> 
                                                        </div>
                                            </div>
                                        </div>
                                    </div>
                                <?php else: ?>
                                    —
                                <?php endif; ?>
                            </td>
                            <td class="text-center">
                                <?php if ($item['order_status'] === 'PENDING'): ?>
                                    <button type="button"
                                        class="btn btn-sm btn-danger btn-delete"
                                        data-id="<?= $item['order_item_id'] ?>">
                                        🗑️
                                    </button>

                                <?php elseif ($item['order_status'] === 'PAID'): ?>
                                    <span class="badge bg-info">Đơn hàng đang chờ xử lý</span>
                                
                                <?php elseif ($item['order_status'] === 'CANCELLED'): ?>
                                    <span class="badge bg-danger">Đơn hàng đã bị hủy</span>

                                <?php elseif (in_array($item['order_status'], ['APPROVED'])): ?>
                                    <span class="badge bg-success">Đơn hàng đã được duyệt, đang ship đến</span>

                                <?php else: ?>
                                    <span class="badge bg-success">Hoàn tất</span>
                                <?php endif; ?>
                                </td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>

            <div class="d-flex justify-content-between align-items-center mt-3">
                <?php if (($_GET['status'] ?? 'PENDING') === 'PENDING'): ?>
                    <h4>Tổng tiền: <span id="total-price">0₫</span></h4>
                    <button type="button" id="checkout-btn" class="btn btn-success">Đặt hàng ngay</button>
                <?php endif; ?>
            </div>
        </form>
    <?php endif; ?>
</div>

<script>
    // Giữ nguyên phần script của bạn, nó đã hoạt động tốt cho AJAX và tính toán
</script>

<?php require_once __DIR__ . '/../layouts/footer.php'; ?>

    <script>
        const selectAll = document.getElementById('select-all');
        const checkboxes = document.querySelectorAll('.item-checkbox');
        const totalPriceEl = document.getElementById('total-price');

        function formatCurrency(num){
            return num.toLocaleString('vi-VN') + '₫';
        }

        function calculateTotal(){
        let total = 0;
        document.querySelectorAll('tr[data-order-item-id]').forEach(tr=>{
            const checkbox = tr.querySelector('.item-checkbox');
            if(checkbox && checkbox.checked){
                const quantity = parseInt(tr.querySelector('.quantity-input').value);
                const price = parseInt(tr.querySelector('.price').textContent.replace(/\D/g,''));
                total += quantity * price;
            }
        });
        totalPriceEl.textContent = formatCurrency(total);
        }


        // Chọn tất cả
        selectAll.addEventListener('change', ()=>{
            checkboxes.forEach(cb => { if(!cb.disabled) cb.checked = selectAll.checked; });
            calculateTotal();
        });

        // Checkbox từng item
        checkboxes.forEach(cb=>{
            cb.addEventListener('change', ()=>{
                selectAll.checked = [...checkboxes].filter(c=>!c.disabled).every(c => c.checked);
                calculateTotal();
            });
        });

        // Nút tăng giảm
        document.querySelectorAll('tr[data-order-item-id]').forEach(tr => {
            const decreaseBtn = tr.querySelector('.btn-decrease');
            const increaseBtn = tr.querySelector('.btn-increase');
            const qtyInput = tr.querySelector('.quantity-input');
            const priceEl = tr.querySelector('.price');
            const subtotalEl = tr.querySelector('.subtotal');

            const price = parseInt(priceEl.textContent.replace(/\D/g, ''));

            if (decreaseBtn && increaseBtn) {
                // 1. Xử lý nút Tăng (+)
                increaseBtn.addEventListener('click', () => {
                    let currentVal = parseInt(qtyInput.value);
                    if (currentVal >= 50) {
                        alert("Số lượng tối đa là 50 sản phẩm");
                        return;
                    }
                    fetch('/GocCaPhe/public/index.php?url=cart/updateQuantity', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: `id=${tr.dataset.orderItemId}&type=inc`
                    })
                    .then(r => r.json())
                    .then(res => {
                        if (res.success) {
                            qtyInput.value = res.quantity;
                            subtotalEl.textContent = formatCurrency(res.quantity * price);
                            calculateTotal();
                        }
                    });
                });

                // 2. Xử lý nút Giảm (-)
                // 2. Xử lý nút Giảm (-)
                decreaseBtn.addEventListener('click', () => {
                    let currentVal = parseInt(qtyInput.value);
    
                    // Bắt lỗi nếu bấm giảm khi đang là 1
                    if (currentVal <= 1) {
                        alert("Số lượng tối thiểu là 1 sản phẩm. Nếu không muốn mua, vui lòng nhấn nút xóa!");
                        return;
                    }

                    fetch('/GocCaPhe/public/index.php?url=cart/updateQuantity', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: `id=${tr.dataset.orderItemId}&type=dec`
                    })
                    .then(r => r.json())
                    .then(res => {
                        if (res.success) {
                            qtyInput.value = res.quantity;
                            subtotalEl.textContent = formatCurrency(res.quantity * price);
                            calculateTotal();
                        }
                    });
                });

                // 3. Xử lý khi NHẬP TRỰC TIẾP bằng bàn phím
                qtyInput.addEventListener('change', () => {
                    let val = parseInt(qtyInput.value);

                    // Kiểm tra nếu vượt quá 50
                    if (val > 50) {
                        alert("Số lượng đã vượt quá giới hạn (tối đa 50). Xin quý khách vui lòng nhập lại!");
                        qtyInput.value = 1; // Reset về 1
                        qtyInput.focus();   // Đưa con trỏ chuột vào lại ô nhập
                        return;             // Thoát ra, không gửi AJAX lên server
                    }

                    // Kiểm tra tính hợp lệ khác
                    if (isNaN(val) || val < 1) {
                        qtyInput.value = 0;
                        return;
                    }

                    // Nếu hợp lệ (1-50) thì mới gửi AJAX
                    fetch('/GocCaPhe/public/index.php?url=cart/updateQuantity', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: `id=${tr.dataset.orderItemId}&type=set&quantity=${val}`
                    })
                    .then(r => r.json())
                    .then(res => {
                        if (res.success) {
                            subtotalEl.textContent = formatCurrency(val * price);
                            calculateTotal();
                        } else {
                            alert(res.message);
                            qtyInput.value = 0; // Reset về 0 nếu server cũng báo lỗi
                        }
                    });
                });
            }
        });
        // Checkout button
        document.getElementById('checkout-btn').addEventListener('click', ()=>{
            const selectedIds = [...document.querySelectorAll('.item-checkbox:checked')].map(cb=>cb.value);
            if(selectedIds.length===0){
                alert('Vui lòng chọn ít nhất 1 sản phẩm');
                return;
            }

            // Redirect sang trang nhập địa chỉ & thanh toán
            window.location.href = '/GocCaPhe/public/index.php?url=cart/checkout&items='+selectedIds.join(',');
        });
        document.addEventListener('click', function(e){
            if(e.target.closest('.btn-delete')){
                const btn = e.target.closest('.btn-delete');
                const id = btn.dataset.id;

                if(!confirm('Bạn có chắc muốn xóa?')) return;

                fetch('/GocCaPhe/public/index.php?url=cart/delete', {
                    method:'POST',
                    headers:{'Content-Type':'application/x-www-form-urlencoded'},
                    body:'id=' + id
                })
                .then(r=>r.json())
                .then(res=>{
                    if(res.success){
                        btn.closest('tr').remove();
                        calculateTotal();
                    } else {
                        alert('Không thể xóa sản phẩm');
                    }
                });
            }
        });



    // Tính tổng ngay khi load
        calculateTotal();

    </script>


    <?php require_once __DIR__ . '/../layouts/footer.php'; ?>
