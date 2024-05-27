(defun c:zea (/ acapp acdoc aclay) 
  (setq acapp (vlax-get-acad-object)
        acdoc (vla-get-activedocument acapp)
        aclay (vla-get-activelayout acdoc)
  )
  (vlax-for layout (vla-get-layouts acdoc) 
    (vla-put-activelayout acdoc layout)
    (if (eq acpaperspace (vla-get-activespace acdoc)) 
      (vla-put-mspace acdoc :vlax-false)
    )
    (vla-zoomextents acapp)
  )
  (vla-put-activelayout acdoc aclay)
  (princ)
)
(vl-load-com)
(princ)