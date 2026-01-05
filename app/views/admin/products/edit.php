<?php require_once __DIR__ . '/../../layouts/admin_header.php'; ?>

<div class="container mt-4">
    <h2>Sửa Product</h2>
    <form method="POST" action="?url=admin/products/update" enctype="multipart/form-data">
        <input type="hidden" name="id" value="<?= $product['id'] ?>">
        <div class="mb-3">
            <label class="form-label">Tên Product</label>
            <input type="text" name="name" class="form-control" value="<?= htmlspecialchars($product['name']) ?>" required>
        </div>
        <div class="mb-3">
            <label class="form-label">Category</label>
            <select name="category_id" class="form-select" required>
                <option value="">-- Chọn Category --</option>
                <?php foreach ($categories as $cat): ?>
                    <option value="<?= $cat['id'] ?>" <?= $cat['id']==$product['category_id']?'selected':'' ?>>
                        <?= htmlspecialchars($cat['name']) ?>
                    </option>
                <?php endforeach ?>
            </select>
        </div>
        <div class="mb-3">
            <label class="form-label">Price</label>
            <input type="number" name="price" class="form-control" value="<?= $product['price'] ?>" required>
        </div>
        <div class="mb-3">
            <label class="form-label">Status</label>
            <select name="status" class="form-select">
                <option value="AVAILABLE" <?= $product['status']=='AVAILABLE'?'selected':'' ?>>AVAILABLE</option>
                <option value="HIDDEN" <?= $product['status']=='HIDDEN'?'selected':'' ?>>HIDDEN</option>
                <option value="SPECIAL" <?= $product['status']=='SPECIAL'?'selected':'' ?>>SPECIAL</option>
            </select>
        </div>
        <div class="mb-3">
            <label class="form-label">Image</label>
            <input type="file" name="image" class="form-control">
            <?php if ($product['image']): ?>
                <img src="/GocCaPhe/public/assets/img/<?= $product['image'] ?>" alt="" style="max-width:100px;margin-top:5px;">
            <?php endif; ?>
        </div>
        <div class="mb-3">
            <label class="form-label">Description</label>
            <textarea name="description" class="form-control"><?= htmlspecialchars($product['description']) ?></textarea>
        </div>
        <button type="submit" class="btn btn-success">💾 Lưu</button>
        <a href="?url=admin/products" class="btn btn-secondary">Hủy</a>
    </form>
</div>
