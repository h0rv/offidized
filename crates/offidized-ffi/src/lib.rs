use std::ffi::CStr;
use std::os::raw::{c_char, c_int};
use std::ptr;
use std::sync::{Mutex, OnceLock};

use offidized_docx::Document;
use offidized_pptx::Presentation;
use offidized_xlsx::{Workbook, XlsxError};

pub const OFFIDIZED_STATUS_OK: c_int = 0;
pub const OFFIDIZED_STATUS_NULL_POINTER: c_int = 1;
pub const OFFIDIZED_STATUS_INVALID_UTF8: c_int = 2;
pub const OFFIDIZED_STATUS_INVALID_ARGUMENT: c_int = 3;
pub const OFFIDIZED_STATUS_OPERATION_FAILED: c_int = 4;

#[repr(C)]
pub struct OffidizedWorkbook {
    inner: Workbook,
}

#[repr(C)]
pub struct OffidizedDocument {
    inner: Document,
}

#[repr(C)]
pub struct OffidizedPresentation {
    inner: Presentation,
}

static LAST_ERROR: OnceLock<Mutex<String>> = OnceLock::new();

fn with_last_error<R>(func: impl FnOnce(&mut String) -> R) -> R {
    let mutex = LAST_ERROR.get_or_init(|| Mutex::new(String::new()));
    let mut guard = match mutex.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };
    func(&mut guard)
}

fn clear_last_error_internal() {
    with_last_error(String::clear);
}

fn success() -> c_int {
    clear_last_error_internal();
    OFFIDIZED_STATUS_OK
}

fn fail(status: c_int, message: impl Into<String>) -> c_int {
    with_last_error(|slot| {
        *slot = message.into();
    });
    status
}

fn require_mut_ptr<'a, T>(pointer: *mut T, name: &str) -> Result<&'a mut T, c_int> {
    if pointer.is_null() {
        return Err(fail(
            OFFIDIZED_STATUS_NULL_POINTER,
            format!("{name} must not be null"),
        ));
    }

    let reference = unsafe {
        // SAFETY: Pointer nullability was checked above and callers own pointer validity.
        pointer.as_mut()
    };

    match reference {
        Some(reference) => Ok(reference),
        None => Err(fail(
            OFFIDIZED_STATUS_NULL_POINTER,
            format!("{name} must not be null"),
        )),
    }
}

fn read_c_string<'a>(pointer: *const c_char, name: &str) -> Result<&'a str, c_int> {
    if pointer.is_null() {
        return Err(fail(
            OFFIDIZED_STATUS_NULL_POINTER,
            format!("{name} must not be null"),
        ));
    }

    let c_str = unsafe {
        // SAFETY: Pointer nullability was checked above. Caller must provide NUL-terminated storage.
        CStr::from_ptr(pointer)
    };

    match c_str.to_str() {
        Ok(value) => Ok(value),
        Err(error) => Err(fail(
            OFFIDIZED_STATUS_INVALID_UTF8,
            format!("{name} must be valid UTF-8: {error}"),
        )),
    }
}

fn write_nul_terminated_bytes(buffer: *mut c_char, buffer_len: usize, bytes: &[u8]) {
    let copy_len = bytes.len().min(buffer_len.saturating_sub(1));
    unsafe {
        // SAFETY: `buffer` is non-null and `copy_len + 1` is within `buffer_len`.
        ptr::copy_nonoverlapping(bytes.as_ptr(), buffer.cast::<u8>(), copy_len);
        *buffer.add(copy_len) = 0;
    }
}

fn drop_ffi_box<T>(pointer: *mut T) {
    unsafe {
        // SAFETY: Pointer must have been allocated by corresponding `*_create`.
        drop(Box::from_raw(pointer));
    }
}

#[no_mangle]
pub extern "C" fn offidized_last_error_length() -> usize {
    with_last_error(|slot| slot.len())
}

#[no_mangle]
pub extern "C" fn offidized_last_error_message(buffer: *mut c_char, buffer_len: usize) -> usize {
    with_last_error(|slot| {
        let bytes = slot.as_bytes();
        let full_len = bytes.len();

        if buffer.is_null() || buffer_len == 0 {
            return full_len;
        }

        write_nul_terminated_bytes(buffer, buffer_len, bytes);

        full_len
    })
}

#[no_mangle]
pub extern "C" fn offidized_clear_last_error() {
    clear_last_error_internal();
}

#[no_mangle]
pub extern "C" fn offidized_workbook_create(out_workbook: *mut *mut OffidizedWorkbook) -> c_int {
    let out_workbook = match require_mut_ptr(out_workbook, "out_workbook") {
        Ok(out_workbook) => out_workbook,
        Err(status) => return status,
    };

    let workbook = Box::new(OffidizedWorkbook {
        inner: Workbook::new(),
    });
    *out_workbook = Box::into_raw(workbook);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_workbook_free(workbook: *mut OffidizedWorkbook) {
    if workbook.is_null() {
        return;
    }

    drop_ffi_box(workbook);
}

#[no_mangle]
pub extern "C" fn offidized_workbook_add_sheet(
    workbook: *mut OffidizedWorkbook,
    sheet_name: *const c_char,
) -> c_int {
    let workbook = match require_mut_ptr(workbook, "workbook") {
        Ok(workbook) => workbook,
        Err(status) => return status,
    };
    let sheet_name = match read_c_string(sheet_name, "sheet_name") {
        Ok(sheet_name) => sheet_name,
        Err(status) => return status,
    };

    workbook.inner.add_sheet(sheet_name);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_workbook_set_cell_string(
    workbook: *mut OffidizedWorkbook,
    sheet_name: *const c_char,
    cell_reference: *const c_char,
    value: *const c_char,
) -> c_int {
    let workbook = match require_mut_ptr(workbook, "workbook") {
        Ok(workbook) => workbook,
        Err(status) => return status,
    };
    let sheet_name = match read_c_string(sheet_name, "sheet_name") {
        Ok(sheet_name) => sheet_name,
        Err(status) => return status,
    };
    let cell_reference = match read_c_string(cell_reference, "cell_reference") {
        Ok(cell_reference) => cell_reference,
        Err(status) => return status,
    };
    let value = match read_c_string(value, "value") {
        Ok(value) => value,
        Err(status) => return status,
    };

    let Some(sheet) = workbook.inner.sheet_mut(sheet_name) else {
        return fail(
            OFFIDIZED_STATUS_INVALID_ARGUMENT,
            format!("sheet `{sheet_name}` does not exist"),
        );
    };

    match sheet.cell_mut(cell_reference) {
        Ok(cell) => {
            cell.set_value(value);
            success()
        }
        Err(XlsxError::InvalidCellReference(_)) => fail(
            OFFIDIZED_STATUS_INVALID_ARGUMENT,
            format!("cell_reference `{cell_reference}` is invalid"),
        ),
        Err(error) => fail(
            OFFIDIZED_STATUS_OPERATION_FAILED,
            format!("failed to set cell `{cell_reference}` on sheet `{sheet_name}`: {error}"),
        ),
    }
}

#[no_mangle]
pub extern "C" fn offidized_workbook_save(
    workbook: *mut OffidizedWorkbook,
    path: *const c_char,
) -> c_int {
    let workbook = match require_mut_ptr(workbook, "workbook") {
        Ok(workbook) => workbook,
        Err(status) => return status,
    };
    let path = match read_c_string(path, "path") {
        Ok(path) => path,
        Err(status) => return status,
    };

    match workbook.inner.save(path) {
        Ok(()) => success(),
        Err(error) => fail(
            OFFIDIZED_STATUS_OPERATION_FAILED,
            format!("failed to save workbook to `{path}`: {error}"),
        ),
    }
}

#[no_mangle]
pub extern "C" fn offidized_document_create(out_document: *mut *mut OffidizedDocument) -> c_int {
    let out_document = match require_mut_ptr(out_document, "out_document") {
        Ok(out_document) => out_document,
        Err(status) => return status,
    };

    let document = Box::new(OffidizedDocument {
        inner: Document::new(),
    });
    *out_document = Box::into_raw(document);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_document_free(document: *mut OffidizedDocument) {
    if document.is_null() {
        return;
    }

    drop_ffi_box(document);
}

#[no_mangle]
pub extern "C" fn offidized_document_add_paragraph(
    document: *mut OffidizedDocument,
    text: *const c_char,
) -> c_int {
    let document = match require_mut_ptr(document, "document") {
        Ok(document) => document,
        Err(status) => return status,
    };
    let text = match read_c_string(text, "text") {
        Ok(text) => text,
        Err(status) => return status,
    };

    document.inner.add_paragraph(text);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_document_add_heading(
    document: *mut OffidizedDocument,
    text: *const c_char,
    level: u8,
) -> c_int {
    let document = match require_mut_ptr(document, "document") {
        Ok(document) => document,
        Err(status) => return status,
    };
    let text = match read_c_string(text, "text") {
        Ok(text) => text,
        Err(status) => return status,
    };

    if !(1..=9).contains(&level) {
        return fail(
            OFFIDIZED_STATUS_INVALID_ARGUMENT,
            "level must be between 1 and 9",
        );
    }

    document.inner.add_heading(text, level);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_document_save(
    document: *mut OffidizedDocument,
    path: *const c_char,
) -> c_int {
    let document = match require_mut_ptr(document, "document") {
        Ok(document) => document,
        Err(status) => return status,
    };
    let path = match read_c_string(path, "path") {
        Ok(path) => path,
        Err(status) => return status,
    };

    match document.inner.save(path) {
        Ok(()) => success(),
        Err(error) => fail(
            OFFIDIZED_STATUS_OPERATION_FAILED,
            format!("failed to save document to `{path}`: {error}"),
        ),
    }
}

#[no_mangle]
pub extern "C" fn offidized_presentation_create(
    out_presentation: *mut *mut OffidizedPresentation,
) -> c_int {
    let out_presentation = match require_mut_ptr(out_presentation, "out_presentation") {
        Ok(out_presentation) => out_presentation,
        Err(status) => return status,
    };

    let presentation = Box::new(OffidizedPresentation {
        inner: Presentation::new(),
    });
    *out_presentation = Box::into_raw(presentation);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_presentation_free(presentation: *mut OffidizedPresentation) {
    if presentation.is_null() {
        return;
    }

    drop_ffi_box(presentation);
}

#[no_mangle]
pub extern "C" fn offidized_presentation_add_slide_with_title(
    presentation: *mut OffidizedPresentation,
    title: *const c_char,
) -> c_int {
    let presentation = match require_mut_ptr(presentation, "presentation") {
        Ok(presentation) => presentation,
        Err(status) => return status,
    };
    let title = match read_c_string(title, "title") {
        Ok(title) => title,
        Err(status) => return status,
    };

    presentation.inner.add_slide_with_title(title);
    success()
}

#[no_mangle]
pub extern "C" fn offidized_presentation_save(
    presentation: *mut OffidizedPresentation,
    path: *const c_char,
) -> c_int {
    let presentation = match require_mut_ptr(presentation, "presentation") {
        Ok(presentation) => presentation,
        Err(status) => return status,
    };
    let path = match read_c_string(path, "path") {
        Ok(path) => path,
        Err(status) => return status,
    };

    match presentation.inner.save(path) {
        Ok(()) => success(),
        Err(error) => fail(
            OFFIDIZED_STATUS_OPERATION_FAILED,
            format!("failed to save presentation to `{path}`: {error}"),
        ),
    }
}
