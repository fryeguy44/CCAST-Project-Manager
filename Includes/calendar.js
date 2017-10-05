$(document).ready(function() {
	$('#calendar').fullCalendar({
		lang: 'en', // Customize the language and localization options for the calendar
		contentHeight: 480, // Will make the calendar's content area a pixel height
		displayEventEnd: false, // Whether or not to display an event's end time text when it is rendered on the calendar
		editable: true, // Determines whether the events on the calendar can be modified
		eventLimit: true, // Limits the number of events displayed on a day
		events: 'includes/calendar.asp',
		firstDay: 1, // The day that each week begins
		fixedWeekCount: true, // Determines the number of weeks displayed in a month view
		header: { // Defines the buttons and title at the top of the calendar
			left: 'prev,next today',
			center: 'title',
			right: 'month,agendaWeek,agendaDay'
		},
		timezone: 'local', // Determines the timezone in which dates throughout the API are parsed and rendered
		selectable: true, // Allows a user to highlight multiple days or timeslots by clicking and dragging
		select: function(start, end) {
			var start = moment(start).format('YYYY-MM-DD') + ' ' + moment(start).format('HH:mm');
			var end = moment(end).format('YYYY-MM-DD') + ' ' + moment(end).format('HH:mm');
			var valid = moment(end).isValid();
			if(end === null || end === 'null' || valid == false)
				end = start;
			$.facebox({ajax: 'includes/calendar.asp?start='+encodeURIComponent(start)+'&end='+encodeURIComponent(end)});
			$('#calendar').fullCalendar('unselect');
		},
		eventClick: function(calEvent, jsEvent, view) {
			$.facebox({ajax: 'includes/calendar.asp?id='+calEvent.id});
		},
		eventDrop: function(event, delta, revertFunc) {
			var start = moment(event.start).format('YYYY-MM-DD') + ' ' + moment(event.start).format('HH:mm');
			var end = moment(event.end).format('YYYY-MM-DD') + ' ' + moment(event.end).format('HH:mm');
			var valid = moment(event.end).isValid();
			if(event.end === null || event.end === 'null' || valid == false)
				end = start;
			$.ajax({
				type: 'post',
				url: 'includes/calendar.asp?id='+event.id+'&start='+start+'&end='+end,
				dataType: 'html',
				success: function(data){
					$('#calendar').fullCalendar('refetchEvents')
				}
			});
		},
		eventResize: function(event, delta, revertFunc) {
			var start = moment(event.start).format('YYYY-MM-DD') + ' ' + moment(event.start).format('HH:mm');
			var end = moment(event.end).format('YYYY-MM-DD') + ' ' + moment(event.end).format('HH:mm');
			var valid = moment(event.end).isValid();
			if(event.end === null || event.end === 'null' || valid == false)
				end = start;
			$.ajax({
				type: 'post',
				url: 'includes/calendar.asp?id='+event.id+'&start='+start+'&end='+end,
				dataType: 'html',
				success: function(data){
					$('#calendar').fullCalendar('refetchEvents')
				}
			});
		}
	});
});

$(document).bind('reveal.facebox', function() {
	if($('[name=EventColor]').length){ $('[name=EventColor]').colorBox(); }
});

function Validation(url,id) {
	var error = 0;
	var button = id.find('button');
	button.prop('disabled', true);
	$('.req', id).each(function () {
		var input = $(this).val();
		var pattern = /^(.|\n)+$/;
		if (!input || !pattern.test(input)) {
			$(this).addClass('error');
			error++;
		}
		else {
			$(this).removeClass('error');
		}
	});
	if (error === 0) {
		Save(url,id);
	} else {
		button.prop('disabled', false);
	}
	return false;
};

function Question(question, url) {
	$.facebox('<p>' + question + '</p><button id="yes">Yes</button><button onclick="$.facebox.close();">No</button>');
	$('#yes').on('click', function() {
		Save(url,$('#facebox .content'));
	});
	return false;
};

function Save(url, id) {
	var button = id.find('button');
	$.ajax({
		type: 'post',
		url: url,
		data: id.serialize(),
		dataType: 'html',
		timeout: 5000,
		cache: false,
		global: false,
		error: function(xhr){
			button.prop('disabled', false).insertAfter('<p>'+xhr.responseText+'</p>');
		},
		success: function(data){
			$('#facebox').css({
				top: $(window).scrollTop() + ($(window).height() / 10),
				left: $(window).width() / 2 - ($('#facebox .popup').outerWidth() / 2)
			});
			$('#calendar').fullCalendar('refetchEvents');
			id.html(data).delay(1000).queue(function() {
				$.facebox.close();
				$(this).dequeue();
			});
		}
	});
};