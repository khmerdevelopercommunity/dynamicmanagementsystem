<?php
// Ensure uploads folder exists upfront
$uploads_dir = __DIR__ . '/uploads';
if (!is_dir($uploads_dir)) {
    mkdir($uploads_dir, 0777, true);
}

// Handle Builder Setup Submission
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action_builder'])) {
    $fields = $_POST['fields'] ?? [];
    $clean_fields = array_values(array_filter(array_map('trim', $fields)));
    $system_name = !empty($_POST['system_name']) ? trim($_POST['system_name']) : 'Management System';

    // Process optional header logo (File OR URL)
    $system_logo = '';
    if (isset($_FILES['system_logo']) && $_FILES['system_logo']['error'] === UPLOAD_ERR_OK) {
        $ext = strtolower(pathinfo($_FILES['system_logo']['name'], PATHINFO_EXTENSION));
        $allowed = ['jpg', 'jpeg', 'png', 'gif', 'webp'];
        if (in_array($ext, $allowed)) {
            $logo_filename = 'logo_' . uniqid() . '.' . $ext;
            if (move_uploaded_file($_FILES['system_logo']['tmp_name'], $uploads_dir . '/' . $logo_filename)) {
                $system_logo = 'uploads/' . $logo_filename;
            }
        }
    } elseif (!empty($_POST['system_logo_url'])) {
        $system_logo = trim($_POST['system_logo_url']);
    }

    // 1. CREATE DATABASE.DB (SQLite3)
    $db_file = __DIR__ . '/database.db';
    if (file_exists($db_file)) {
        unlink($db_file);
    }
    
    $db = new SQLite3($db_file);
    $db->exec('PRAGMA encoding = "UTF-8";');
    $db->exec('PRAGMA foreign_keys = ON;');

    $columns = [
        'id INTEGER PRIMARY KEY AUTOINCREMENT',
        'avatar TEXT',
        'created_at DATETIME DEFAULT CURRENT_TIMESTAMP'
    ];
    
    foreach ($clean_fields as $field) {
        $columns[] = '"' . str_replace('"', '""', $field) . '" TEXT';
    }

    $query = 'CREATE TABLE IF NOT EXISTS records (' . implode(', ', $columns) . ');';
    $db->exec($query);
    $db->close();

    // 2. GENERATE INDEX.PHP
    $schema_export = var_export($clean_fields, true);
    $logo_export = var_export($system_logo, true);
    $name_export = var_export($system_name, true);

    $index_code = '<?php
header("Content-Type: text/html; charset=utf-8");
$db = new SQLite3(__DIR__ . "/database.db");
$db->exec("PRAGMA encoding = \"UTF-8\";");

$SYSTEM_NAME = ' . $name_export . ';
$SCHEMA_FIELDS = ' . $schema_export . ';
$SYSTEM_LOGO = ' . $logo_export . ';
$error = "";

// Auto-create table if missing
$cols = [
    \'id INTEGER PRIMARY KEY AUTOINCREMENT\',
    \'avatar TEXT\',
    \'created_at DATETIME DEFAULT CURRENT_TIMESTAMP\'
];
foreach ($SCHEMA_FIELDS as $field) {
    $cols[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'" TEXT\';
}
$db->exec(\'CREATE TABLE IF NOT EXISTS records (\' . implode(\', \', $cols) . \');\');

// Helper to remove local server files safely
function removeImageFile($path) {
    if (!empty($path) && strpos($path, "uploads/") === 0) {
        $full_path = __DIR__ . "/" . $path;
        if (file_exists($full_path) && is_file($full_path)) {
            unlink($full_path);
        }
    }
}

// ---------------- Handle EXCEL (.xls) EXPORT ----------------
if (isset($_GET["action"]) && $_GET["action"] === "export") {
    // Sanitize system name for safe filename usage
    $clean_sys_name = preg_replace("/[^a-zA-Z0-9_\-]/", "_", $SYSTEM_NAME);
    $clean_sys_name = trim(preg_replace("/_+/", "_", $clean_sys_name), "_");
    if (empty($clean_sys_name)) {
        $clean_sys_name = "System";
    }

    $filename = $clean_sys_name . "_export_" . date("Y-m-d_H-i-s") . ".xls";
    
    header("Content-Type: application/vnd.ms-excel; charset=utf-8");
    header("Content-Disposition: attachment; filename=" . $filename);
    header("Pragma: no-cache");
    header("Expires: 0");

    echo \'<?xml version="1.0" encoding="UTF-8"?>\' . "\n";
    echo \'<?mso-application progid="Excel.Sheet"?>\' . "\n";
    ?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
              xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:x="urn:schemas-microsoft-com:office:excel"
              xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
      <Styles>
        <Style ss:ID="Header">
          <Font ss:Bold="1" ss:Color="#FFFFFF"/>
          <Interior ss:Color="#4F81BD" ss:Pattern="Solid"/>
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
        </Style>
        <Style ss:ID="Data">
          <Alignment ss:Vertical="Center"/>
        </Style>
      </Styles>
      <Worksheet ss:Name="Records">
        <Table>
          <Row ss:StyleID="Header">
            <Cell><Data ss:Type="String">ID</Data></Cell>
            <?php foreach ($SCHEMA_FIELDS as $field): ?>
              <Cell><Data ss:Type="String"><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></Data></Cell>
            <?php endforeach; ?>
            <Cell><Data ss:Type="String">Profile Image Link / Path</Data></Cell>
            <Cell><Data ss:Type="String">Created At</Data></Cell>
          </Row>
          <?php
          $results = $db->query("SELECT * FROM records ORDER BY id ASC");
          while ($row = $results->fetchArray(SQLITE3_ASSOC)):
              $img_display = $row["avatar"] ?? "";
              $is_url = (strpos($img_display, "http://") === 0 || strpos($img_display, "https://") === 0);
          ?>
            <Row ss:StyleID="Data">
              <Cell><Data ss:Type="Number"><?= $row["id"] ?></Data></Cell>
              <?php foreach ($SCHEMA_FIELDS as $field): ?>
                <Cell><Data ss:Type="String"><?= htmlspecialchars($row[$field] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
              <?php endforeach; ?>
              
              <?php if ($is_url): ?>
                <Cell ss:HRef="<?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?>">
                    <Data ss:Type="String"><?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?></Data>
                </Cell>
              <?php else: ?>
                <Cell><Data ss:Type="String"><?= htmlspecialchars($img_display, ENT_QUOTES, "UTF-8") ?></Data></Cell>
              <?php endif; ?>

              <Cell><Data ss:Type="String"><?= htmlspecialchars($row["created_at"] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
            </Row>
          <?php endwhile; ?>
        </Table>
      </Worksheet>
    </Workbook>
    <?php
    exit;
}

// ---------------- Handle DELETE ----------------
if (isset($_GET["action"]) && $_GET["action"] === "delete" && isset($_GET["id"])) {
    $id = (int)$_GET["id"];
    
    $stmt = $db->prepare("SELECT avatar FROM records WHERE id = :id");
    $stmt->bindValue(":id", $id, SQLITE3_INTEGER);
    $res = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
    
    if ($res) {
        removeImageFile($res["avatar"]);
        
        $del_stmt = $db->prepare("DELETE FROM records WHERE id = :id");
        $del_stmt->bindValue(":id", $id, SQLITE3_INTEGER);
        $del_stmt->execute();
    }
    header("Location: index.php");
    exit;
}

// ---------------- Handle CREATE & UPDATE ----------------
if ($_SERVER["REQUEST_METHOD"] === "POST" && isset($_POST["action"])) {
    $action = $_POST["action"];
    $record_id = isset($_POST["id"]) ? (int)$_POST["id"] : 0;
    $avatar_path = "";
    $allowed = ["jpg", "jpeg", "png", "gif", "webp"];

    // Fetch existing record if updating
    $existing = null;
    if ($action === "update" && $record_id > 0) {
        $stmt = $db->prepare("SELECT * FROM records WHERE id = :id");
        $stmt->bindValue(":id", $record_id, SQLITE3_INTEGER);
        $existing = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
        $avatar_path = $existing["avatar"] ?? "";
    }

    $target_dir = __DIR__ . "/uploads/";
    if (!is_dir($target_dir)) {
        mkdir($target_dir, 0777, true);
    }

    // 1. Check direct file upload (PC/Phone)
    if (isset($_FILES["avatar"]) && $_FILES["avatar"]["error"] === UPLOAD_ERR_OK) {
        $tmp_name = $_FILES["avatar"]["tmp_name"];
        $ext = strtolower(pathinfo($_FILES["avatar"]["name"], PATHINFO_EXTENSION));
        
        if (in_array($ext, $allowed)) {
            $filename = uniqid("img_", true) . "." . $ext;
            if (move_uploaded_file($tmp_name, $target_dir . $filename)) {
                if ($action === "update" && !empty($existing["avatar"])) {
                    removeImageFile($existing["avatar"]);
                }
                $avatar_path = "uploads/" . $filename;
            } else {
                $error = "Failed to upload file to local server.";
            }
        } else {
            $error = "Invalid file format! Allowed: JPG, JPEG, PNG, GIF, WEBP.";
        }
    } 
    // 2. Direct Web Image URL
    elseif (!empty($_POST["avatar_url"])) {
        $url = trim($_POST["avatar_url"]);
        if (filter_var($url, FILTER_VALIDATE_URL)) {
            if ($action === "update" && !empty($existing["avatar"])) {
                removeImageFile($existing["avatar"]);
            }
            $avatar_path = $url;
        } else {
            $error = "Invalid Image URL provided.";
        }
    } 
    elseif ($action === "create" && empty($avatar_path)) {
        $error = "Please upload an image file or provide an Image URL.";
    }

    if (empty($error)) {
        if ($action === "create") {
            $cols = ["avatar"];
            $vals = [":avatar"];
            $params = [":avatar" => $avatar_path];

            foreach ($SCHEMA_FIELDS as $field) {
                $cols[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'"\';
                $param_key = ":" . md5($field);
                $vals[] = $param_key;
                $params[$param_key] = $_POST[$field] ?? "";
            }

            $stmt = $db->prepare("INSERT INTO records (" . implode(", ", $cols) . ") VALUES (" . implode(", ", $vals) . ")");
            foreach ($params as $key => $val) {
                $stmt->bindValue($key, $val, SQLITE3_TEXT);
            }
            $stmt->execute();
        } 
        elseif ($action === "update" && $record_id > 0) {
            $set_clauses = ["avatar = :avatar"];
            $params = [":avatar" => $avatar_path, ":id" => $record_id];

            foreach ($SCHEMA_FIELDS as $field) {
                $param_key = ":" . md5($field);
                $set_clauses[] = \'"\' . str_replace(\'"\', \'""\', $field) . \'" = \' . $param_key;
                $params[$param_key] = $_POST[$field] ?? "";
            }

            $stmt = $db->prepare("UPDATE records SET " . implode(", ", $set_clauses) . " WHERE id = :id");
            foreach ($params as $key => $val) {
                $type = ($key === ":id") ? SQLITE3_INTEGER : SQLITE3_TEXT;
                $stmt->bindValue($key, $val, $type);
            }
            $stmt->execute();
        }

        header("Location: index.php");
        exit;
    }
}

// Fetch record to edit
$edit_data = null;
if (isset($_GET["action"]) && $_GET["action"] === "edit" && isset($_GET["id"])) {
    $edit_id = (int)$_GET["id"];
    $stmt = $db->prepare("SELECT * FROM records WHERE id = :id");
    $stmt->bindValue(":id", $edit_id, SQLITE3_INTEGER);
    $edit_data = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
}

$results = $db->query("SELECT * FROM records ORDER BY id ASC");
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><?= htmlspecialchars($SYSTEM_NAME, ENT_QUOTES, "UTF-8") ?></title>
    <style>
        body { font-family: sans-serif; background: #f0f2f5; margin: 0; padding: 20px; }
        .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
        .header-left { display: flex; align-items: center; gap: 15px; }
        .header img { height: 50px; border-radius: 6px; }
        .main-layout { display: flex; gap: 20px; align-items: flex-start; }
        .card-form { flex: 1; min-width: 320px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .card-table { flex: 2; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); min-width: 0; }
        .table-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input[type="text"], input[type="file"], input[type="url"] { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
        .upload-row { display: flex; gap: 10px; align-items: center; }
        .upload-row input[type="file"], .upload-row input[type="url"] { flex: 1; }
        .preview-avatar-box { width: 70px; height: 70px; border-radius: 50%; object-fit: cover; border: 2px solid #ddd; margin-bottom: 10px; display: block; }
        .btn-submit { background: #28a745; color: white; border: none; padding: 10px; width: 100%; border-radius: 4px; cursor: pointer; font-size: 16px; }
        .btn-cancel { display: block; text-align: center; background: #6c757d; color: white; text-decoration: none; padding: 8px; border-radius: 4px; margin-top: 10px; }
        .btn-export { background: #217346; color: white; text-decoration: none; padding: 8px 14px; border-radius: 4px; font-weight: bold; font-size: 14px; white-space: nowrap; }
        .error-msg { background: #f8d7da; color: #721c24; padding: 10px; border-radius: 4px; margin-bottom: 15px; }
        
        /* ONLY HORIZONTAL SCROLL FOR TABLE */
        .table-responsive { 
            width: 100%; 
            overflow-x: auto; 
            overflow-y: visible; 
            -webkit-overflow-scrolling: touch; 
            border: 1px solid #e2e8f0; 
            border-radius: 6px; 
        }
        table { width: 100%; border-collapse: separate; border-spacing: 0; min-width: max-content; }
        th, td { border-bottom: 1px solid #ddd; border-right: 1px solid #ddd; padding: 10px; text-align: left; white-space: nowrap; min-width: 130px; }
        th:last-child, td:last-child { border-right: none; }
        th { background: #f8f9fa; }
        
        .avatar-img { width: 45px; height: 45px; object-fit: cover; border-radius: 50%; display: block; }
        .action-btn { padding: 4px 8px; text-decoration: none; border-radius: 4px; font-size: 12px; margin-right: 4px; display: inline-block; }
        .btn-edit { background: #ffc107; color: #000; }
        .btn-delete { background: #dc3545; color: #fff; }

        @media (max-width: 768px) {
            .main-layout { flex-direction: column; }
            .card-form, .card-table { width: 100%; min-width: 100%; box-sizing: border-box; }
        }
    </style>
</head>
<body>

<div class="header">
    <div class="header-left">
        <?php if (!empty($SYSTEM_LOGO)): ?>
            <?php 
                $logo_src = (strpos($SYSTEM_LOGO, "http://") === 0 || strpos($SYSTEM_LOGO, "https://") === 0) 
                    ? $SYSTEM_LOGO 
                    : htmlspecialchars($SYSTEM_LOGO, ENT_QUOTES, "UTF-8");
            ?>
            <img src="<?= $logo_src ?>" alt="Logo">
        <?php endif; ?>
        <h2><?= htmlspecialchars($SYSTEM_NAME, ENT_QUOTES, "UTF-8") ?></h2>
    </div>
</div>

<?php if (!empty($error)): ?>
    <div class="error-msg"><?= htmlspecialchars($error, ENT_QUOTES, "UTF-8") ?></div>
<?php endif; ?>

<div class="main-layout">
    <div class="card-form">
        <h3><?= $edit_data ? "Edit Record" : "Add New Record" ?></h3>
        <form method="POST" enctype="multipart/form-data">
            <input type="hidden" name="action" value="<?= $edit_data ? "update" : "create" ?>">
            <?php if ($edit_data): ?>
                <input type="hidden" name="id" value="<?= $edit_data["id"] ?>">
            <?php endif; ?>

            <div class="form-group">
                <label>Profile Image Preview</label>
                <?php 
                $initial_img = "data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' width=\'70\' height=\'70\' viewBox=\'0 0 24 24\' fill=\'%23ccc\'><path d=\'M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z\'/></svg>";
                if ($edit_data && !empty($edit_data["avatar"])) {
                    $initial_img = htmlspecialchars($edit_data["avatar"], ENT_QUOTES, "UTF-8");
                }
                ?>
                <img id="avatar-preview" src="<?= $initial_img ?>" class="preview-avatar-box" alt="Avatar Preview">

                <div class="upload-row">
                    <input type="file" name="avatar" accept="image/*" onchange="previewAvatarFile(this)">
                    <input type="url" id="avatar_url_input" name="avatar_url" placeholder="or Image URL (https://...)" oninput="previewAvatarUrl(this.value)" value="<?= ($edit_data && strpos($edit_data[\'avatar\'], \'http\') === 0) ? htmlspecialchars($edit_data[\'avatar\'], ENT_QUOTES, \'UTF-8\') : \'\' ?>">
                </div>
            </div>

            <?php foreach ($SCHEMA_FIELDS as $field): ?>
            <div class="form-group">
                <label><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></label>
                <input type="text" 
                       name="<?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?>" 
                       value="<?= htmlspecialchars($edit_data[$field] ?? "", ENT_QUOTES, "UTF-8") ?>" 
                       required>
            </div>
            <?php endforeach; ?>

            <button type="submit" class="btn-submit"><?= $edit_data ? "Update Record" : "Save Record" ?></button>
            <?php if ($edit_data): ?>
                <a href="index.php" class="btn-cancel">Cancel Edit</a>
            <?php endif; ?>
        </form>
    </div>
    
    <div class="card-table">
        <div class="table-header">
            <h3 style="margin:0;">System Records</h3>
            <a href="index.php?action=export" class="btn-export">📊 Export (.XLS Excel)</a>
        </div>
        
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th style="min-width: 65px;">Avatar</th>
                        <?php foreach ($SCHEMA_FIELDS as $field): ?>
                        <th><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></th>
                        <?php endforeach; ?>
                        <th style="min-width: 110px;">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <?php while ($row = $results->fetchArray(SQLITE3_ASSOC)): ?>
                    <tr>
                        <td>
                            <?php if (!empty($row["avatar"])): ?>
                                <img src="<?= htmlspecialchars($row["avatar"], ENT_QUOTES, "UTF-8") ?>" class="avatar-img">
                            <?php else: ?>
                                <span>No Image</span>
                            <?php endif; ?>
                        </td>
                        <?php foreach ($SCHEMA_FIELDS as $field): ?>
                        <td><?= htmlspecialchars($row[$field] ?? "", ENT_QUOTES, "UTF-8") ?></td>
                        <?php endforeach; ?>
                        <td>
                            <a href="index.php?action=edit&id=<?= $row["id"] ?>" class="action-btn btn-edit">Edit</a>
                            <a href="index.php?action=delete&id=<?= $row["id"] ?>" class="action-btn btn-delete" onclick="return confirm(\'Are you sure you want to delete this record?\')">Delete</a>
                        </td>
                    </tr>
                    <?php endwhile; ?>
                </tbody>
            </table>
        </div>
    </div>
</div>

<script>
function previewAvatarFile(input) {
    const preview = document.getElementById("avatar-preview");
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            preview.src = e.target.result;
        }
        reader.readAsDataURL(input.files[0]);
    }
}

function previewAvatarUrl(url) {
    const preview = document.getElementById("avatar-preview");
    if (url.trim() !== "") {
        preview.src = url;
    }
}
</script>

</body>
</html>';

    file_put_contents(__DIR__ . '/index.php', $index_code);

    header('Location: index.php');
    exit;
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Dynamic System Builder</title>
    <style>
        body { font-family: sans-serif; background: #f4f6f8; padding: 30px; }
        .container { max-width: 850px; margin: auto; background: #fff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .builder-grid { display: flex; gap: 20px; margin-top: 20px; }
        .left-col { flex: 1; border: 2px dashed #bbb; padding: 20px; border-radius: 6px; background: #fafafa; }
        .right-col { flex: 2; }
        .form-group { margin-bottom: 15px; }
        label { display: block; font-weight: bold; margin-bottom: 5px; }
        input[type="text"], input[type="file"], input[type="url"] { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        .field-row { display: flex; gap: 8px; margin-bottom: 10px; }
        .preview-img { max-width: 100%; max-height: 100px; margin-top: 10px; display: none; border-radius: 4px; border: 1px solid #ddd; }
        .btn-add { background: #17a2b8; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; }
        .btn-remove { background: #dc3545; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; }
        .btn-done { background: #28a745; color: white; border: none; width: 100%; padding: 14px; font-size: 18px; font-weight: bold; border-radius: 4px; margin-top: 20px; cursor: pointer; }
    </style>
</head>
<body>
<div class="container">
    <h2>Dynamic System Builder</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="hidden" name="action_builder" value="1">
        
        <div class="form-group">
            <label>System / Project Name</label>
            <input type="text" name="system_name" placeholder="e.g. Employee Directory, Student Management System, Warehouse Inventory" required>
        </div>

        <div class="builder-grid">
            <div class="left-col">
                <h3>System Logo</h3>
                <p style="color:#666; font-size: 13px;">Optional logo for header banner.</p>
                
                <div style="margin-bottom: 10px;">
                    <input type="file" name="system_logo" accept="image/*" onchange="previewImageFile(this)">
                </div>
                <div style="margin-bottom: 10px;">
                    <input type="url" name="system_logo_url" placeholder="or Logo URL (https://...)" oninput="previewImageUrl(this.value)">
                </div>
                
                <img id="logo-preview" class="preview-img" alt="Logo Preview">
            </div>

            <div class="right-col">
                <h3>Custom Dynamic Fields</h3>
                <div id="dynamic-fields">
                    <div class="field-row">
                        <input type="text" name="fields[]" placeholder="Field Name (e.g. Full Name, Age, Position)" required>
                        <button type="button" class="btn-remove" onclick="removeRow(this)">Remove</button>
                    </div>
                </div>
                <button type="button" class="btn-add" onclick="addRow()">+ Add Field</button>
            </div>
        </div>
        <button type="submit" class="btn-done">Done (Build Database & Index)</button>
    </form>
</div>

<script>
function previewImageFile(input) {
    const preview = document.getElementById('logo-preview');
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            preview.src = e.target.result;
            preview.style.display = 'block';
        }
        reader.readAsDataURL(input.files[0]);
    }
}

function previewImageUrl(url) {
    const preview = document.getElementById('logo-preview');
    if (url.trim() !== '') {
        preview.src = url;
        preview.style.display = 'block';
    }
}

function addRow() {
    const container = document.getElementById('dynamic-fields');
    const div = document.createElement('div');
    div.className = 'field-row';
    div.innerHTML = `
        <input type="text" name="fields[]" placeholder="Field Name (Unicode Supported)" required>
        <button type="button" class="btn-remove" onclick="removeRow(this)">Remove</button>
    `;
    container.appendChild(div);
}

function removeRow(btn) {
    const container = document.getElementById('dynamic-fields');
    if (container.children.length > 1) {
        btn.parentElement.remove();
    }
}
</script>
</body>
</html>