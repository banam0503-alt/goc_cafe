<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>

<div class="container mt-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h2>Quản lý Products</h2>
        <a href="?url=admin/products/create" class="btn btn-primary">➕ Thêm Product</a>
    </div>

    <table class="table table-striped table-bordered align-middle">
        <thead class="table-dark">
    <tr>
        <th>ID</th>
        <th>Tên Product</th>
        <th>Category</th>
        <th>Price</th>
        <th>Description</th>
        <th>Status</th>
        <th>Ảnh</th> <!-- cột mới -->
        <th class="text-center">Hành động</th>
    </tr>
</thead>
<tbody>
<?php foreach ($products as $p): ?>
    <tr>
        <td><?= $p['id'] ?></td>
        <td><?= htmlspecialchars($p['name']) ?></td>
        <td><?= htmlspecialchars($p['category_name'] ?? 'Không có') ?></td>
        <td><?= number_format($p['price']) ?>₫</td>
        <td><?= htmlspecialchars($p['description']) ?></td>
        <td>
    <?php
        $statusClass = 'bg-secondary';
        if ($p['status'] === 'AVAILABLE') {
            $statusClass = 'bg-success';
        } elseif ($p['status'] === 'SPECIAL') {
            $statusClass = 'bg-warning';
        }
    ?>
    <span class="badge <?= $statusClass ?>">
        <?= $p['status'] ?>
    </span>
</td>

        <td class="text-center">
            <?php if($p['image']): ?>
                <!-- Button mở modal -->
                <button type="button" class="btn btn-sm btn-info" data-bs-toggle="modal" data-bs-target="#imgModal<?= $p['id'] ?>">
                    👁️
                </button>

                <!-- Modal Bootstrap -->
                <div class="modal fade" id="imgModal<?= $p['id'] ?>" tabindex="-1" aria-labelledby="imgModalLabel<?= $p['id'] ?>" aria-hidden="true">
                  <div class="modal-dialog modal-dialog-centered">
                    <div class="modal-content">
                      <div class="modal-header">
                        <h5 class="modal-title" id="imgModalLabel<?= $p['id'] ?>"><?= htmlspecialchars($p['name']) ?></h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                      </div>
                      <div class="modal-body text-center">
                        <img src="/GocCaPhe/public/assets/img/<?= $p['image'] ?>" alt="<?= htmlspecialchars($p['name']) ?>" class="img-fluid">
                      </div>
                    </div>
                  </div>
                </div>
            <?php else: ?>
                —
            <?php endif; ?>
        </td>
        <td class="text-center">
            <a href="?url=admin/products/edit&id=<?= $p['id'] ?>" class="btn btn-sm btn-success me-1">✏️ Sửa</a>
            <a href="?url=admin/products/delete&id=<?= $p['id'] ?>" class="btn btn-sm btn-danger" onclick="return confirm('Bạn có chắc muốn xóa product này?')">🗑️ Xóa</a>
        </td>
    </tr>
<?php endforeach ?>
</tbody>
    </table>
</div>
