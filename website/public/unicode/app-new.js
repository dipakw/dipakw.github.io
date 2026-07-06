$(document).ready(function(){
	/**
	 * Tabbing
	 */
	$( document ).on( 'click', '.tab', function(){

		if( ! $(this).hasClass('active') ) {
			cc( $(this) );
		}

		var index = $(this).attr('index');
		$('[index]').removeClass('active');
		$('[index='+index+']').addClass('active');
	} );
	$('.tab').each(function(i){
		$(this).attr( 'index', i );
	});
	$('.inputs').each(function(i){
		$(this).attr( 'index', i );
	});
	/**
	 * Add copy button
	 */
	$( '.input2 .control' ).each(function( i ){
		var control = $(this);
		var report = i + 1;
		control.append('<span report="' + report + '" class="copy">COPY</span>');
	});
	function cc( element ) {

		//return false;

		var id = element.attr('report') || '';
		if ( id == '' || isNaN( id ) ) {
			return false;
		}
		
		var action = ['', 'Unicode', 'Preeti'][id];
		gtag('event', 'Click', { 'event_category': 'Interactive', 'event_label': 'Switch Unicode', 'non_interaction': true })
		return false;

		var img = $('#cc');
		var link =  "http://xtlook.com/unicode/ccount/click.php?id=" + id;
		link += '&rand=' + $.now();
		img.attr( 'src', link );
	}
	/**
	 * Copy
	 */
	var bin = document.getElementById('bin');
	$( document ).on('click','.copy',function(){

		var textarea = $(this).parent('.control').next('textarea');
		var value = textarea.val();
		if ( value == '' ){
			return false;
		}
		bin.value = value;
		bin.focus();
		bin.select();
		document.execCommand('copy');

		textarea.fadeOut();
		setTimeout(function(){
			textarea.fadeIn();
		},1200);

		cc( $(this) );

	});

	$(document).on('keyup','#u2p_unicode',function(){
		$('#u2p_preeti').val( xtl_u2p( $(this).val() ) );
	});
	
});