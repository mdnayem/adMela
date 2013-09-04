<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<script language="javascript">
function ShowReg(op) {
  document.getElementById('Car').style.display='none';
  document.getElementById('Boat').style.display='none';
  document.getElementById('Plane').style.display='none';

  if (op == '1') {
    document.getElementById('Car').style.display="block";
  }
  if (op == 2) {
    document.getElementById('Boat').style.display="block";
  }
  if (op == 3) {
    document.getElementById('Plane').style.display="block";
  }
}

</script>

<select id="choice" onChange="ShowReg(this.selectedIndex)" size="4">
  <option value='0'>Select registration type</option>
  <option value="1">Car</option>
  <option value="2">Boat</option>
  <option value="3">Plane</option>
</select>
<br>
<div id="Car" style="display:none">
  VIN No. <input type="text" id="VIN" value="">
</div>
<div id="Boat" style="display:none">
  Registration No. <input type="text" id="Breg" value="">
</div>
<div id="Plane" style="display:none">
  Tail No. <input type="text" id="Ptail" value="">
  Model <input type="text" id="Pmodel" value="">
</div>


</body>
</html>
