document.addEventListener('DOMContentLoaded', function() {
	
	//GO TO TOP::
	var viewport = $('html, body');
	//animate to top:
	$('.gototop').click(function(){
    	 viewport.animate({scrollTop : 0}, 1400, 'easeOutQuint');
    });
    //stop animation on scroll
    viewport.bind("scroll mousedown DOMMouseScroll mousewheel keyup", function(e){
	     viewport.stop();
    });

    //Tooltips init:
    if ($('[data-toggle="tooltip"]').length > 0) {
        $('[data-toggle="tooltip"]').tooltip();
    }
    //New init by attribute
    if ($('[tooltip]').length > 0) {
        $('[tooltip]').tooltip({
            trigger: 'hover',
            title: function () {
                return $(this).attr("tooltip-title");
            }
        });
    }
	
	//sections slide Up/Down
    $("h2.section-heading, h3.section-heading, h4.section-heading").click(function(){
		section = $(this).data('section');
		if($(this).hasClass("opened")){
			$("section[data-section="+section+"]").slideUp(450);
		}
		else{
			$("section[data-section="+section+"]").slideDown(450);
		}
		$(this).toggleClass("opened");
	});
	
    //bootstrap confirmation
    if ($('[data-toggle=confirmation]').length != 0) {
        $('[data-toggle=confirmation]').confirmation({
            container: 'body',
            //placement: 'top',
            //btnOkLabel: 'Да',
            //btnCancelLabel: 'Не',
            popout: true
        });
    }
    //New init by attribute
    if ($('[confirmation]').length != 0) {
        $('[confirmation]').confirmation({
            container: 'body',
            title: function () {
                return $(this).attr("confirm-title");
            },
            //placement: 'top',
            //btnOkLabel: 'Да',
            //btnCancelLabel: 'Не',
            popout: true
        });
    }

    //bootstrap popover
	if ($('[data-toggle=popover]').length != 0) {
	    $('[data-toggle="popover"]').popover();
	}

    //edit disabled form
	if ($('input[data-action=edit]').length != 0) {
	    $('input[data-action=edit]').click(function (e) {
	        if ($(this).attr('data-action') == 'edit') {
	            e.preventDefault();
	            $(this).removeAttr('data-action').attr('value', $(this).attr('data-save-label')).closest('form').find('[disabled]').removeAttr('disabled');
	        }
	    });
	}

    //table filters toggle
	if ($('.filter-toggle').length != 0) {
	    $('.filter-toggle').click(function () {
	        if ($(this).hasClass('opened')) {
	            $(this).parent().find("form").slideUp();
	        }
	        else {
	            $(this).parent().find("form").slideDown();
	        }
	        $(this).toggleClass('opened');
	    });
	}

	//status histories toggle
	if ($('.toggle').length != 0) {
		$('.toggle').click(function () {
			if ($(this).hasClass('opened')) {
				$(this).parent().parent().next().slideUp();
			}
			else {
				$(this).parent().parent().next().slideDown();
			}
			$(this).toggleClass('opened');
		});
	}

    //table filters sum and date validation
	if($('.table-filters form').length != 0){
	    //if not empty => change comma to dot
	    //set number to 2 decimals
	    //if field is not a number => set empty value
	    $(".table-filters form input[name='SearchDO.PrAmountFrom'], .table-filters form input[name='SearchDO.PrAmountTo'], .table-filters form input[name='SearchDO.TransactionAmountFrom'], .table-filters form input[name='SearchDO.TransactionAmountTo']").bind('change', function(event) {
	        if($(this).val() != ""){
	            res = $(this).val().replace(",", ".");
	            res = parseFloat(res).toFixed(2);
	            if(res != 'NaN' && res > 0 && res <= 99999) $(this).val(res);
	            else $(this).val("");
	        }
			
	    });
	}

    //Datepicker
	if($('.datepicker').length > 0){
	    $('input.datepicker').datepicker({
	        format: "dd.mm.yyyy",
	        language: "bg",
	        autoclose: true,
	        keyboardNavigation: false,
	        forceParse: false
	    });
	    //if format != 'dd.mm.yyyy' => set empty value:
	    $('input.datepicker').change(function(){
	        if (!(/^(((0[1-9]|[12]\d|3[01])\.(0[13578]|1[02])\.((19|[2-9]\d)\d{2}))|((0[1-9]|[12]\d|30)\.(0[13456789]|1[012])\.((19|[2-9]\d)\d{2}))|((0[1-9]|1\d|2[0-8])\.02\.((19|[2-9]\d)\d{2}))|(29\.02\.((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))))$/.test($(this).val()))) {
	            $(this).val("");
	        }
	    });
	}

    //Form reset button
	if($('form .reset').length > 0){
	    $('form .reset').click(function(){
	        form = $(this).closest('form');
	        form.find("input[type!=submit], textarea").val("");
	        form.find("select option").removeAttr("selected");	//IE type
	        form.submit()
	    });
	}

});