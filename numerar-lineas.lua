-- Numerar todas las lineas de todos los bloques de codigo.
-- numberLines es la clase nativa de pandoc (skylighting): en el HTML emite
-- un ancla <a> con aria-label por linea, y el numero visible lo pinta
-- plantilla-pdf.css con ::before (el filtro no mete texto en el codigo).
function CodeBlock (cb)
  cb.classes:insert('numberLines')
  return cb
end
