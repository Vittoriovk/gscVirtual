	<script src="/gscVirtual/bootstrap/js/jquery.min.js"></script>
	<script src="/gscVirtual/bootstrap/js/popper.min.js"></script>
	<script src="/gscVirtual/bootstrap/js/bootstrap.min.js"></script>
	<script src="/gscVirtual/js/function.js"></script>
	<script src="/gscVirtual/js/bootbox.min.js"></script>
	<script src="/gscVirtual/js/chosen.jquery.js" type="text/javascript"></script>

    <script src="/gscVirtual/js/ab-datepicker-2.1.17/js/locales/it.min.js" type="text/javascript"></script>
    <script src="/gscVirtual/js/ab-datepicker-2.1.17/js/datepicker.min.js" type="text/javascript"></script>
	
<script>
	$('.mydatepicker').datepicker({
		inputFormat: ["dd/MM/yyy"],
		outputFormat: 'dd/MM/yyyy'
	});
</script>

<script>
$(document).ready(function(){
   $(".chosen").chosen({ width:'100%' });
});

</script>