sub iterateProjects()
  
  dim projs as vbProjects
  dim proj  as vbProject
  dim refs  as References
  dim ref   as Reference
  dim comps as vbComponents
  dim comp  as vbComponent
    
  dim i as integer
  dim r as integer
  dim c as integer
    
  set projs = application.VBE.vbProjects
    
  for i = 1 to projs.Count
  
      set proj = projs(i)
      debug.print "Project: " & proj.Name ' default: "VBAProject"
      debug.print "  description:       " & proj.description
      debug.print "  buildFileName:     " & proj.buildFileName
      debug.print "  mode:              " & proj.mode
      
      set refs = proj.References
      debug.print ""      
      debug.print "  References"
      for r = 1 to refs.Count
          set ref = refs(r)
          debug.print "    Reference: " & ref.Name
          debug.print "      Guid: "    & ref.GUID
          debug.print "      path: "    & ref.fullPath
          debug.print "      version: " & ref.major & "." & ref.minor
          debug.print "      type: "    & ref.type
      next r
      
      debug.print "  Components"
      set comps = proj.VBComponents
      for c = 1 to comps.Count
          set comp = comps(c)
          debug.print "    Component: "     & comp.name
          debug.print "      type:      "   & comp.type
          debug.print "      code module: " & comp.codeModule
      next c
  
  next i
  
end sub

