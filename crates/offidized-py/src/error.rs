use offidized_docx::DocxError;
use offidized_opc::OpcError;
use offidized_pptx::PptxError;
use offidized_xlsx::XlsxError;
use pyo3::create_exception;
use pyo3::exceptions::{
    PyException, PyIOError, PyNotImplementedError, PyRuntimeError, PyValueError,
};
use pyo3::prelude::*;

create_exception!(offidized, OffidizedError, PyException);
create_exception!(offidized, OffidizedIoError, PyIOError);
create_exception!(offidized, OffidizedValueError, PyValueError);
create_exception!(offidized, OffidizedUnsupportedError, PyNotImplementedError);
create_exception!(offidized, OffidizedRuntimeError, PyRuntimeError);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExceptionKind {
    Io,
    Value,
    Unsupported,
    Runtime,
}

pub(crate) fn register_exceptions(py: Python<'_>, module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add("OffidizedError", py.get_type::<OffidizedError>())?;
    module.add("OffidizedIoError", py.get_type::<OffidizedIoError>())?;
    module.add("OffidizedValueError", py.get_type::<OffidizedValueError>())?;
    module.add(
        "OffidizedUnsupportedError",
        py.get_type::<OffidizedUnsupportedError>(),
    )?;
    module.add(
        "OffidizedRuntimeError",
        py.get_type::<OffidizedRuntimeError>(),
    )?;
    Ok(())
}

pub(crate) fn value_error(message: impl Into<String>) -> PyErr {
    build_py_err(ExceptionKind::Value, message.into())
}

pub(crate) fn xlsx_error_to_py(error: XlsxError) -> PyErr {
    build_py_err(classify_xlsx_error(&error), error.to_string())
}

pub(crate) fn docx_error_to_py(error: DocxError) -> PyErr {
    build_py_err(classify_docx_error(&error), error.to_string())
}

pub(crate) fn pptx_error_to_py(error: PptxError) -> PyErr {
    build_py_err(classify_pptx_error(&error), error.to_string())
}

fn build_py_err(kind: ExceptionKind, message: String) -> PyErr {
    match kind {
        ExceptionKind::Io => OffidizedIoError::new_err(message),
        ExceptionKind::Value => OffidizedValueError::new_err(message),
        ExceptionKind::Unsupported => OffidizedUnsupportedError::new_err(message),
        ExceptionKind::Runtime => OffidizedRuntimeError::new_err(message),
    }
}

fn classify_xlsx_error(error: &XlsxError) -> ExceptionKind {
    match error {
        XlsxError::Io(_) => ExceptionKind::Io,
        XlsxError::Opc(opc) => classify_opc_error(opc),
        XlsxError::InvalidCellReference(_)
        | XlsxError::InvalidWorkbookState(_)
        | XlsxError::InvalidFormula(_) => ExceptionKind::Value,
        XlsxError::UnsupportedPackage(_) => ExceptionKind::Unsupported,
        XlsxError::Utf8(_)
        | XlsxError::Xml(_)
        | XlsxError::XmlSerialize(_)
        | XlsxError::XmlDeserialize(_) => ExceptionKind::Runtime,
    }
}

fn classify_docx_error(error: &DocxError) -> ExceptionKind {
    match error {
        DocxError::Io(_) => ExceptionKind::Io,
        DocxError::Opc(opc) => classify_opc_error(opc),
        DocxError::UnsupportedPackage(_) => ExceptionKind::Unsupported,
        DocxError::Xml(_) => ExceptionKind::Runtime,
    }
}

fn classify_pptx_error(error: &PptxError) -> ExceptionKind {
    match error {
        PptxError::Io(_) => ExceptionKind::Io,
        PptxError::Opc(opc) => classify_opc_error(opc),
        PptxError::UnsupportedPackage(_) => ExceptionKind::Unsupported,
        PptxError::Xml(_) => ExceptionKind::Runtime,
        PptxError::InvalidOperation(_) => ExceptionKind::Value,
    }
}

fn classify_opc_error(error: &OpcError) -> ExceptionKind {
    match error {
        OpcError::Io(_) | OpcError::Zip(_) => ExceptionKind::Io,
        OpcError::PartNotFound(_)
        | OpcError::InvalidRelationship(_)
        | OpcError::InvalidContentType(_)
        | OpcError::InvalidUri(_) => ExceptionKind::Value,
        OpcError::Xml(_) | OpcError::MalformedPackage(_) => ExceptionKind::Runtime,
    }
}
