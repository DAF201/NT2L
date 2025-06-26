import os
import core.REGEX
import core.TOOL
import service.pod_build
import config.config
import config.regex_reference
import functions.target_table
import functions.image_process


def rename_pods(file_paths: list[str]) -> list[str]:
    """for each pdf, try to read data and rename, return failed files"""
    core.TOOL.message("start building POD")

    target_table_excel = functions.target_table.target_table()

    attention_needed = set()
    file_foler = os.path.abspath(file_paths[0]).replace(
        os.path.basename(file_paths[0]), "")
    for file in file_paths:
        try:
            core.TOOL.message(f"collecting information from file: {file}")
            images = functions.image_process.pdf_to_images(file)
            pdf_content = functions.image_process.read_images(images)

            FXSJ_pgi = core.REGEX.re_find_first(
                config.regex_reference.OCR_FXSJ_PGI_REGEX, pdf_content)
            pb = core.REGEX.re_find_first(
                config.regex_reference.OCR_FXSJ_PB_REGEX, pdf_content)
            FXSJ_amount = core.REGEX.re_find_first(
                config.regex_reference.OCR_FXSJ_AMOUNT_REGEX, pdf_content)
            regular_dn = core.REGEX.re_find_first(
                config.regex_reference.OCR_DN_REGEX, pdf_content)

            if FXSJ_pgi:
                # FXSJ vendor pool units
                core.TOOL.message("FXSJ POD found")
                os.rename(file, os.path.join(
                    file_foler, f"{pb}_{FXSJ_amount}x_FXSJ_{core.TOOL.get_today_date_pad2()}"))
                continue
            if regular_dn:
                # regular pod
                core.TOOL.message("ordinary POD found")
                total_amount = 0
                for recipient_data in core.REGEX.re_find_all(config.regex_reference.OCR_RECIPIENT_REGEX, pdf_content):
                    total_amount += int(recipient_data[6])

                destination = target_table_excel.search_dn_for_info(regular_dn[-8:])[
                    "Ship to"]
                os.rename(file, os.path.join(
                    file_foler, f"{regular_dn[-8:]}_{pb}_{total_amount}x_{destination}_{core.TOOL.get_today_date_pad2()}.pdf"))
                continue
            core.TOOL.message("cannot determine POD type")
            attention_needed.add(file)
        except:
            attention_needed.add(file)
    core.TOOL.message("POD building complete")
    return list(attention_needed)


def build_pod() -> list[str] | None:
    """download pdf, read pdf, rename pdf, return the list of pod need to be done manually"""
    pdf_files = service.pod_build.download_pods(
        config.config.global_config["POD"]["save_path"])
    if len(pdf_files) == 0:
        return None
    res = rename_pods(pdf_files)
    return res if res != [] else None
