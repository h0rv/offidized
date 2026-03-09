#ifndef OFFIDIZED_FFI_H
#define OFFIDIZED_FFI_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef struct OffidizedWorkbook OffidizedWorkbook;
typedef struct OffidizedDocument OffidizedDocument;
typedef struct OffidizedPresentation OffidizedPresentation;

typedef enum OffidizedStatus {
    OFFIDIZED_STATUS_OK = 0,
    OFFIDIZED_STATUS_NULL_POINTER = 1,
    OFFIDIZED_STATUS_INVALID_UTF8 = 2,
    OFFIDIZED_STATUS_INVALID_ARGUMENT = 3,
    OFFIDIZED_STATUS_OPERATION_FAILED = 4,
} OffidizedStatus;

size_t offidized_last_error_length(void);
size_t offidized_last_error_message(char* buffer, size_t buffer_len);
void offidized_clear_last_error(void);

int32_t offidized_workbook_create(OffidizedWorkbook** out_workbook);
void offidized_workbook_free(OffidizedWorkbook* workbook);
int32_t offidized_workbook_add_sheet(OffidizedWorkbook* workbook, const char* sheet_name);
int32_t offidized_workbook_set_cell_string(
    OffidizedWorkbook* workbook,
    const char* sheet_name,
    const char* cell_reference,
    const char* value
);
int32_t offidized_workbook_save(OffidizedWorkbook* workbook, const char* path);

int32_t offidized_document_create(OffidizedDocument** out_document);
void offidized_document_free(OffidizedDocument* document);
int32_t offidized_document_add_paragraph(OffidizedDocument* document, const char* text);
int32_t offidized_document_add_heading(OffidizedDocument* document, const char* text, uint8_t level);
int32_t offidized_document_save(OffidizedDocument* document, const char* path);

int32_t offidized_presentation_create(OffidizedPresentation** out_presentation);
void offidized_presentation_free(OffidizedPresentation* presentation);
int32_t offidized_presentation_add_slide_with_title(
    OffidizedPresentation* presentation,
    const char* title
);
int32_t offidized_presentation_save(OffidizedPresentation* presentation, const char* path);

#ifdef __cplusplus
}
#endif

#endif
