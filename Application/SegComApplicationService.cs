using System;
using System.Collections.Generic;
using XNSeguimentCompres.Data;
using XNSeguimentCompres.Domain;

namespace XNSeguimentCompres.Application
{
    /// <summary>
    /// Coordina la operación de guardado de seguimiento.
    /// UI → Domain (validación) → Repository (persistencia)
    /// </summary>
    public class SegComApplicationService
    {
        private readonly SegComValidator _validator;
        private readonly SegComRepository _repository;

        public SegComApplicationService(
            SegComValidator validator,
            SegComRepository repository)
        {
            _validator = validator;
            _repository = repository;
        }

        /// <summary>
        /// Intenta guardar un nuevo seguimiento.
        /// Devuelve true si tuvo éxito, junto con el nuevo DocEntry.
        /// </summary>
        public bool TryAdd(SegComHeader header, List<SegComLine> lines, out int newDocEntry, out string error)
        {
            // Validación
            if (!_validator.Validate(header, lines, out error))
            {
                newDocEntry = 0;
                return false;
            }

            // Guardar
            newDocEntry = _repository.Add(header, lines);
            return true;
        }

        public bool TryUpdate(SegComHeader header, List<SegComLine> lines, out string error)
        {
            if (!_validator.Validate(header, lines, out error))
                return false;

            _repository.Update(header, lines);
            return true;
        }

        public bool TryGet(int docEntry, out SegComDocument document, out string error)
        {
            document = null;
            try
            {
                document = _repository.GetByDocEntry(docEntry);
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public bool TryDelete(int docEntry, out string error)
        {
            try
            {
                _repository.Delete(docEntry);
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

    }
}

