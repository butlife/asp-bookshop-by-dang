//<![CDATA[
$(function(){
	(function(){
		var curr = 0;
		$("#focusNav .trigger").each(function(i){
			$(this).click(function(){
				curr = i;
				$("#focus img").eq(i).fadeIn("slow").siblings("img").hide();
				$(this).siblings(".trigger").removeClass("imgSelected").end().addClass("imgSelected");
				return false;
			});
		});
		
		var pg = function(flag){
			//flag:true��ʾǰ���� false��ʾ��
			if (flag) {
				if (curr == 0) {
					todo = 2;
				} else {
					todo = (curr - 1) % 3;
				}
			} else {
				todo = (curr + 1) % 3;
			}
			$("#focusNav .trigger").eq(todo).click();
		};
		
		//ǰ��
		$("#prev").click(function(){
			pg(true);
			return false;
		});
		
		//��
		$("#next").click(function(){
			pg(false);
			return false;
		});
		
		//�Զ���
		var timer = setInterval(function(){
			todo = (curr + 1) % 3;
			$("#focusNav .trigger").eq(todo).click();
		},3000);
		
		//�����ͣ�ڴ�������ʱֹͣ�Զ���
		$("#focusNav a").hover(function(){
				clearInterval(timer);
			},
			function(){
				timer = setInterval(function(){
					todo = (curr + 1) % 3;
					$("#focusNav .trigger").eq(todo).click();
				},4000);			
			}
		);
	})();
});
//]]>