<!DOCTYPE html>
<html>
<body>

<form action="/upload" method="post" enctype="multipart/form-data">
    {{ csrf_field() }}    
    Select image to upload:
    <input type="file" name="marksheet" id="fileToUpload">
    <input type="submit" value="Upload Image" name="submit">
</form>

</body>
</html>