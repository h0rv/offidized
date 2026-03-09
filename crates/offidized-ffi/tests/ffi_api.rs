use std::ffi::{CStr, CString};
use std::os::raw::c_char;
use std::path::Path;
use std::ptr;
use std::sync::{Mutex, MutexGuard, OnceLock};

use offidized_ffi::*;
use tempfile::tempdir;

static TEST_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

fn cstring(input: &str) -> CString {
    CString::new(input).expect("input must not contain interior NUL")
}

fn test_lock() -> MutexGuard<'static, ()> {
    let mutex = TEST_MUTEX.get_or_init(|| Mutex::new(()));
    match mutex.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    }
}

fn last_error_string() -> String {
    let len = offidized_last_error_length();
    let mut buffer = vec![0u8; len + 1];
    let full_len = offidized_last_error_message(buffer.as_mut_ptr().cast::<c_char>(), buffer.len());
    assert_eq!(full_len, len);
    let cstr = CStr::from_bytes_until_nul(&buffer).expect("error buffer should be NUL terminated");
    cstr.to_str()
        .expect("last-error should always be valid UTF-8")
        .to_owned()
}

fn assert_file_written(path: &Path) {
    assert!(
        path.exists(),
        "expected file to exist at {}",
        path.display()
    );
    let size = std::fs::metadata(path)
        .expect("metadata should be available")
        .len();
    assert!(size > 0, "expected non-empty file at {}", path.display());
}

#[test]
fn workbook_happy_path_and_save() {
    let _guard = test_lock();
    offidized_clear_last_error();

    let mut workbook = ptr::null_mut();
    assert_eq!(
        offidized_workbook_create(&mut workbook),
        OFFIDIZED_STATUS_OK
    );
    assert!(!workbook.is_null());

    let sheet_name = cstring("Sheet1");
    assert_eq!(
        offidized_workbook_add_sheet(workbook, sheet_name.as_ptr()),
        OFFIDIZED_STATUS_OK
    );

    let cell_ref = cstring("A1");
    let value = cstring("hello ffi");
    assert_eq!(
        offidized_workbook_set_cell_string(
            workbook,
            sheet_name.as_ptr(),
            cell_ref.as_ptr(),
            value.as_ptr()
        ),
        OFFIDIZED_STATUS_OK
    );

    let dir = tempdir().expect("tempdir should be created");
    let path = dir.path().join("workbook.xlsx");
    let path_c = cstring(path.to_str().expect("temp path should be UTF-8"));
    assert_eq!(
        offidized_workbook_save(workbook, path_c.as_ptr()),
        OFFIDIZED_STATUS_OK
    );
    assert_file_written(&path);

    assert_eq!(offidized_last_error_length(), 0);
    offidized_workbook_free(workbook);
}

#[test]
fn document_and_presentation_happy_path_and_save() {
    let _guard = test_lock();
    offidized_clear_last_error();

    let mut document = ptr::null_mut();
    assert_eq!(
        offidized_document_create(&mut document),
        OFFIDIZED_STATUS_OK
    );
    assert!(!document.is_null());

    let heading = cstring("Title");
    let paragraph = cstring("Paragraph body");
    assert_eq!(
        offidized_document_add_heading(document, heading.as_ptr(), 1),
        OFFIDIZED_STATUS_OK
    );
    assert_eq!(
        offidized_document_add_paragraph(document, paragraph.as_ptr()),
        OFFIDIZED_STATUS_OK
    );

    let dir = tempdir().expect("tempdir should be created");
    let doc_path = dir.path().join("doc.docx");
    let doc_path_c = cstring(doc_path.to_str().expect("temp path should be UTF-8"));
    assert_eq!(
        offidized_document_save(document, doc_path_c.as_ptr()),
        OFFIDIZED_STATUS_OK
    );
    assert_file_written(&doc_path);
    offidized_document_free(document);

    let mut presentation = ptr::null_mut();
    assert_eq!(
        offidized_presentation_create(&mut presentation),
        OFFIDIZED_STATUS_OK
    );
    assert!(!presentation.is_null());

    let title = cstring("Slide One");
    assert_eq!(
        offidized_presentation_add_slide_with_title(presentation, title.as_ptr()),
        OFFIDIZED_STATUS_OK
    );

    let pptx_path = dir.path().join("slides.pptx");
    let pptx_path_c = cstring(pptx_path.to_str().expect("temp path should be UTF-8"));
    assert_eq!(
        offidized_presentation_save(presentation, pptx_path_c.as_ptr()),
        OFFIDIZED_STATUS_OK
    );
    assert_file_written(&pptx_path);
    offidized_presentation_free(presentation);

    assert_eq!(offidized_last_error_length(), 0);
}

#[test]
fn null_pointer_and_invalid_argument_paths() {
    let _guard = test_lock();
    offidized_clear_last_error();

    assert_eq!(
        offidized_workbook_create(ptr::null_mut()),
        OFFIDIZED_STATUS_NULL_POINTER
    );
    assert!(last_error_string().contains("out_workbook must not be null"));

    let mut workbook = ptr::null_mut();
    assert_eq!(
        offidized_workbook_create(&mut workbook),
        OFFIDIZED_STATUS_OK
    );

    let missing_sheet = cstring("Missing");
    let cell_ref = cstring("A1");
    let value = cstring("x");
    assert_eq!(
        offidized_workbook_set_cell_string(
            workbook,
            missing_sheet.as_ptr(),
            cell_ref.as_ptr(),
            value.as_ptr()
        ),
        OFFIDIZED_STATUS_INVALID_ARGUMENT
    );
    assert!(last_error_string().contains("does not exist"));

    let sheet_name = cstring("Sheet1");
    assert_eq!(
        offidized_workbook_add_sheet(workbook, sheet_name.as_ptr()),
        OFFIDIZED_STATUS_OK
    );

    let bad_cell_ref = cstring("A0");
    assert_eq!(
        offidized_workbook_set_cell_string(
            workbook,
            sheet_name.as_ptr(),
            bad_cell_ref.as_ptr(),
            value.as_ptr()
        ),
        OFFIDIZED_STATUS_INVALID_ARGUMENT
    );
    assert!(last_error_string().contains("cell_reference `A0` is invalid"));

    offidized_workbook_free(workbook);

    let mut document = ptr::null_mut();
    assert_eq!(
        offidized_document_create(&mut document),
        OFFIDIZED_STATUS_OK
    );
    let title = cstring("Bad heading");
    assert_eq!(
        offidized_document_add_heading(document, title.as_ptr(), 0),
        OFFIDIZED_STATUS_INVALID_ARGUMENT
    );
    assert!(last_error_string().contains("level must be between 1 and 9"));
    offidized_document_free(document);
}

#[test]
fn invalid_utf8_and_last_error_copy_behavior() {
    let _guard = test_lock();
    offidized_clear_last_error();

    let mut workbook = ptr::null_mut();
    assert_eq!(
        offidized_workbook_create(&mut workbook),
        OFFIDIZED_STATUS_OK
    );

    let invalid_utf8 = CString::from_vec_with_nul(vec![0xFF, 0])
        .expect("byte vector should be valid C string storage");
    assert_eq!(
        offidized_workbook_add_sheet(workbook, invalid_utf8.as_ptr()),
        OFFIDIZED_STATUS_INVALID_UTF8
    );

    let full_len = offidized_last_error_length();
    assert!(full_len > 0);

    let mut tiny = vec![0u8; 4];
    let reported = offidized_last_error_message(tiny.as_mut_ptr().cast::<c_char>(), tiny.len());
    assert_eq!(reported, full_len);
    assert_eq!(tiny[tiny.len() - 1], 0);

    let null_buffer = [0u8; 1];
    let len_without_copy = offidized_last_error_message(ptr::null_mut(), null_buffer.len());
    assert_eq!(len_without_copy, full_len);

    offidized_clear_last_error();
    assert_eq!(offidized_last_error_length(), 0);

    offidized_workbook_free(workbook);
}
