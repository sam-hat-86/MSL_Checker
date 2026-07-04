// script.js
// Simple accessibility helpers: smooth scroll and focus management for in-page links

/* サイト内で使う外部リンクや表示ラベルをここで一元管理します。
   例: `vmcDownload` のキーを index.html の data-link-key / data-label-key と対応させる */
// Only configure URLs here. Labels will be generated from the URL automatically.
const SITE_LINKS = {
	vmcDownload: 'https://github.com/sam-hat-86/VMC-VMoshiCode/raw/refs/heads/main/VMC-VMoshiCode2605.xlsm'
};

document.addEventListener('DOMContentLoaded', function(){
	// data-link-key / data-label-key を参照してリンクを自動設定
	(function applySiteLinks(){
		document.querySelectorAll('[data-link-key]').forEach(function(a){
			var key = a.getAttribute('data-link-key');
			if(!key) return;
			var url = SITE_LINKS[key];
			if(url){
				a.setAttribute('href', url);
				a.setAttribute('rel', 'noopener');
				// set a readable label inside the link if a span exists
				var span = a.querySelector('span');
				if(span){
					span.textContent = filenameFromUrl(url)+"をダウンロード";
				}
				// also update aria-label if present on inner container
				var container = a.querySelector('[role="link"]');
				if(container){
					container.setAttribute('aria-label', filenameFromUrl(url)+"をダウンロード");
				}
			}
		});
		// helper: extract filename-like label from URL
		function filenameFromUrl(url){
			try{
				var u = String(url).split('?')[0].split('#')[0];
				var parts = u.split('/');
				var last = parts.pop() || parts.pop(); // handle trailing slash
				last = decodeURIComponent(last || '');
				// If the last segment contains an extension, return it as-is (keep filename.ext)
				if (/\.[^/.]+$/.test(last)){
					return last || url;
				}
				// otherwise, keep hyphens and underscores as-is, only collapse whitespace
				var name = last.replace(/\s+/g,' ').trim();
				return name || url;
			}catch(e){ return url; }
		}

		// set labels for elements that want a generated label
		document.querySelectorAll('[data-label-key]').forEach(function(el){
			var key = el.getAttribute('data-label-key');
			if(!key) return;
			var url = SITE_LINKS[key];
			if(url){
				el.textContent = filenameFromUrl(url)+"をダウンロード";
			}
		});
	})();
	// Smooth scroll for same-page anchors
	document.querySelectorAll('a[href^="#"]').forEach(function(link){
		link.addEventListener('click', function(e){
			var targetId = this.getAttribute('href').slice(1);
			var target = document.getElementById(targetId);
			if(target){
				e.preventDefault();
				target.scrollIntoView({behavior:'smooth', block:'start'});
				// make focus visible for keyboard users
				target.setAttribute('tabindex','-1');
				target.focus({preventScroll:true});
			}
		});
	});
	// No responsive nav toggle needed (TOC is inline)

	// Scrollspy: highlight current section in nav
	var sections = document.querySelectorAll('main section[id]');
	var navLinks = document.querySelectorAll('.toc a');
	function onScroll(){
		var fromTop = window.scrollY + 100;
		var currentId = null;
		sections.forEach(function(section){
			if(section.offsetTop <= fromTop && (section.offsetTop + section.offsetHeight) > fromTop){
				currentId = section.id;
			}
		});
		navLinks.forEach(function(link){
			var href = link.getAttribute('href').slice(1);
			if(href === currentId){
				link.classList.add('active');
				link.setAttribute('aria-current','true');
			} else {
				link.classList.remove('active');
				link.removeAttribute('aria-current');
			}
		});
	}
	window.addEventListener('scroll', onScroll, {passive:true});
	onScroll();
});