;; 1. 添加自定义菜单（正确宏：cx-add-menu，参数=菜单标识+助记符）
(cx-add-menu "CFD Tools" #\p)  ;; #\p是助记符（Alt+P激活菜单）

;; 2. 定义菜单项回调函数（保存Case/Data并启动ParaView）
(define (save_and_launch_paraview)
    (define base-name (strip-directory (in-package cl-file-package rc-filename)))
    (define cas-file (string-append base-name ".cas.h5"))
    (define dat-file (string-append base-name ".dat.h5")) 
    ;; 保存Case和Data文件（TUI命令）
    (ti-menu-load-string (format #f "file/write-case-data ~a ok" cas-file))
    (newline)
    (display (format #f "Saved: ~a, ~a" cas-file dat-file))
    (newline)
    
    ;; 检查cas及dat文件是否存在，若存在则直接启动带参数的ParaView，否则只启动ParaView
    ;; 在后台启动ParaView（不等待进程结束）
    (if (and (file-exists? cas-file) (file-exists? dat-file))
        (begin      
            (display (format #f "File exist: ~a and ~a , launch ParaView..." cas-file dat-file))  
            (newline)
            (system (format #f "powershell.exe -NoProfile -Command \"Start-Process paraview.exe -ArgumentList '--data=\"~a\"' \"" cas-file))
        )
        (begin
            (display (format #f "File not exist: ~a or ~a, launch ParaView without data file..." cas-file dat-file)) 
            (newline)
            (system "powershell.exe -NoProfile -Command \"Start-Process paraview.exe\"") 
        )      
    )
)


;; 3. 添加菜单项
(cx-add-item "CFD Tools" 
             "Launch ParaView" 
             #f #f #t
             save_and_launch_paraview
             )  
