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

    // Process optional header logo
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

    $query = 'CREATE TABLE IF NOT EXISTS students (' . implode(', ', $columns) . ');';
    $db->exec($query);
    $db->close();

    // 2. GENERATE INDEX.PHP WITH HORIZONTAL SCROLLBAR FOR WIDE TABLES
    $schema_export = var_export($clean_fields, true);
    $logo_export = var_export($system_logo, true);

    $index_code = '<?php
header("Content-Type: text/html; charset=utf-8");
$db = new SQLite3(__DIR__ . "/database.db");
$db->exec("PRAGMA encoding = \"UTF-8\";");

$SCHEMA_FIELDS = ' . $schema_export . ';
$SYSTEM_LOGO = ' . $logo_export . ';
$error = "";

// Helper to remove image files from server safely
function removeImageFile($path) {
    if (!empty($path)) {
        $full_path = __DIR__ . "/" . $path;
        if (file_exists($full_path) && is_file($full_path)) {
            unlink($full_path);
        }
    }
}

// ---------------- Handle EXCEL (.xls) EXPORT ----------------
if (isset($_GET["action"]) && $_GET["action"] === "export") {
    $filename = "students_export_" . date("Y-m-d_H-i-s") . ".xls";
    
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
      <Worksheet ss:Name="Students">
        <Table>
          <Row ss:StyleID="Header">
            <Cell><Data ss:Type="String">ID</Data></Cell>
            <?php foreach ($SCHEMA_FIELDS as $field): ?>
              <Cell><Data ss:Type="String"><?= htmlspecialchars($field, ENT_QUOTES, "UTF-8") ?></Data></Cell>
            <?php endforeach; ?>
            <Cell><Data ss:Type="String">Profile Image Path</Data></Cell>
            <Cell><Data ss:Type="String">Created At</Data></Cell>
          </Row>
          <?php
          $results = $db->query("SELECT * FROM students ORDER BY id ASC");
          while ($row = $results->fetchArray(SQLITE3_ASSOC)):
          ?>
            <Row ss:StyleID="Data">
              <Cell><Data ss:Type="Number"><?= $row["id"] ?></Data></Cell>
              <?php foreach ($SCHEMA_FIELDS as $field): ?>
                <Cell><Data ss:Type="String"><?= htmlspecialchars($row[$field] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
              <?php endforeach; ?>
              <Cell><Data ss:Type="String"><?= htmlspecialchars($row["avatar"] ?? "", ENT_QUOTES, "UTF-8") ?></Data></Cell>
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
    
    $stmt = $db->prepare("SELECT avatar FROM students WHERE id = :id");
    $stmt->bindValue(":id", $id, SQLITE3_INTEGER);
    $res = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
    
    if ($res) {
        removeImageFile($res["avatar"]);
        
        $del_stmt = $db->prepare("DELETE FROM students WHERE id = :id");
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

    // Fetch existing record if updating
    $existing = null;
    if ($action === "update" && $record_id > 0) {
        $stmt = $db->prepare("SELECT * FROM students WHERE id = :id");
        $stmt->bindValue(":id", $record_id, SQLITE3_INTEGER);
        $existing = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
        $avatar_path = $existing["avatar"] ?? "";
    }

    // Process new image upload if provided
    if (isset($_FILES["avatar"]) && $_FILES["avatar"]["error"] === UPLOAD_ERR_OK) {
        $tmp_name = $_FILES["avatar"]["tmp_name"];
        $ext = strtolower(pathinfo($_FILES["avatar"]["name"], PATHINFO_EXTENSION));
        $allowed = ["jpg", "jpeg", "png", "gif", "webp"];
        
        if (in_array($ext, $allowed)) {
            $target_dir = __DIR__ . "/uploads/";
            if (!is_dir($target_dir)) {
                mkdir($target_dir, 0777, true);
            }
            
            $filename = uniqid("img_", true) . "." . $ext;
            if (move_uploaded_file($tmp_name, $target_dir . $filename)) {
                if ($action === "update" && !empty($existing["avatar"])) {
                    removeImageFile($existing["avatar"]);
                }
                $avatar_path = "uploads/" . $filename;
            } else {
                $error = "Failed to upload image. Check permissions.";
            }
        } else {
            $error = "Invalid format! Only JPG, JPEG, PNG, GIF, and WEBP allowed.";
        }
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

            $stmt = $db->prepare("INSERT INTO students (" . implode(", ", $cols) . ") VALUES (" . implode(", ", $vals) . ")");
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

            $stmt = $db->prepare("UPDATE students SET " . implode(", ", $set_clauses) . " WHERE id = :id");
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
    $stmt = $db->prepare("SELECT * FROM students WHERE id = :id");
    $stmt->bindValue(":id", $edit_id, SQLITE3_INTEGER);
    $edit_data = $stmt->execute()->fetchArray(SQLITE3_ASSOC);
}

$results = $db->query("SELECT * FROM students ORDER BY id ASC");
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Student Management System</title>
    <style>
        body { font-family: sans-serif; background: #f0f2f5; margin: 0; padding: 20px; }
        .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
        .header-left { display: flex; align-items: center; gap: 15px; }
        .header img { height: 50px; border-radius: 6px; }
        .main-layout { display: flex; gap: 20px; }
        .card-form { flex: 1; min-width: 320px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .card-table { flex: 2; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); min-width: 0; }
        .table-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input[type="text"], input[type="file"] { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
        .btn-submit { background: #28a745; color: white; border: none; padding: 10px; width: 100%; border-radius: 4px; cursor: pointer; font-size: 16px; }
        .btn-cancel { display: block; text-align: center; background: #6c757d; color: white; text-decoration: none; padding: 8px; border-radius: 4px; margin-top: 10px; }
        .btn-export { background: #217346; color: white; text-decoration: none; padding: 8px 14px; border-radius: 4px; font-weight: bold; font-size: 14px; white-space: nowrap; }
        .error-msg { background: #f8d7da; color: #721c24; padding: 10px; border-radius: 4px; margin-bottom: 15px; }
        
        /* Table Scroll Container */
        .table-responsive { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; min-width: max-content; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; white-space: nowrap; min-width: 140px; }
        th { background: #f8f9fa; position: sticky; top: 0; }
        
        .avatar-img { width: 45px; height: 45px; object-fit: cover; border-radius: 50%; display: block; }
        .action-btn { padding: 4px 8px; text-decoration: none; border-radius: 4px; font-size: 12px; margin-right: 4px; display: inline-block; }
        .btn-edit { background: #ffc107; color: #000; }
        .btn-delete { background: #dc3545; color: #fff; }
    </style>
</head>
<body>

<div class="header">
    <div class="header-left">
        <?php if (!empty($SYSTEM_LOGO) && file_exists(__DIR__ . "/" . $SYSTEM_LOGO)): ?>
            <img src="<?= htmlspecialchars($SYSTEM_LOGO, ENT_QUOTES, "UTF-8") ?>" alt="Logo">
        <?php endif; ?>
        <h2>Student Management System</h2>
    </div>
</div>

<?php if (!empty($error)): ?>
    <div class="error-msg"><?= htmlspecialchars($error, ENT_QUOTES, "UTF-8") ?></div>
<?php endif; ?>

<div class="main-layout">
    <div class="card-form">
        <h3><?= $edit_data ? "Edit Student Record" : "Add New Student" ?></h3>
        <form method="POST" enctype="multipart/form-data">
            <input type="hidden" name="action" value="<?= $edit_data ? "update" : "create" ?>">
            <?php if ($edit_data): ?>
                <input type="hidden" name="id" value="<?= $edit_data["id"] ?>">
            <?php endif; ?>

            <div class="form-group">
                <label>Profile Image <?= $edit_data ? "(Leave empty to keep current)" : "" ?></label>
                <?php if ($edit_data && !empty($edit_data["avatar"])): ?>
                    <img src="<?= htmlspecialchars($edit_data["avatar"], ENT_QUOTES, "UTF-8") ?>" class="avatar-img" style="margin-bottom:10px;">
                <?php endif; ?>
                <input type="file" name="avatar" accept="image/*" <?= $edit_data ? "" : "required" ?>>
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

            <button type="submit" class="btn-submit"><?= $edit_data ? "Update Student Record" : "Save Record" ?></button>
            <?php if ($edit_data): ?>
                <a href="index.php" class="btn-cancel">Cancel Edit</a>
            <?php endif; ?>
        </form>
    </div>
    
    <div class="card-table">
        <div class="table-header">
            <h3 style="margin:0;">Student Records</h3>
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
                            <?php if (!empty($row["avatar"]) && file_exists(__DIR__ . "/" . $row["avatar"])): ?>
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
    <title>Create Index & Database</title>
    <style>
        body { font-family: sans-serif; background: #f4f6f8; padding: 30px; }
        .container { max-width: 850px; margin: auto; background: #fff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .builder-grid { display: flex; gap: 20px; margin-top: 20px; }
        .left-col { flex: 1; border: 2px dashed #bbb; padding: 20px; text-align: center; border-radius: 6px; background: #fafafa; }
        .right-col { flex: 2; }
        .field-row { display: flex; gap: 8px; margin-bottom: 10px; }
        input[type="text"], input[type="file"] { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        .preview-img { max-width: 100%; max-height: 120px; margin-top: 10px; display: none; border-radius: 4px; border: 1px solid #ddd; }
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
        <div class="builder-grid">
            <div class="left-col">
                <h3>System Logo / Header</h3>
                <p style="color:#666; font-size: 13px;">Optional logo for your system header.</p>
                <input type="file" name="system_logo" accept="image/*" onchange="previewImage(this)">
                <img id="logo-preview" class="preview-img" alt="Logo Preview">
            </div>

            <div class="right-col">
                <h3>Custom Dynamic Fields</h3>
                <div id="dynamic-fields">
                    <div class="field-row">
                        <input type="text" name="fields[]" placeholder="Field Name (e.g. ឈ្មោះ, Full Name, 年齡)" required>
                        <button type="button" class="btn-remove" onclick="removeRow(this)">Remove</button>
                    </div>
                </div>
                <button type="button" class="btn-add" onclick="addRow()">+ Add Field</button>
            </div>
        </div>
        <button type="submit" class="btn-done">Done (Build database.db & index.php)</button>
    </form>
</div>

<script>
function previewImage(input) {
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