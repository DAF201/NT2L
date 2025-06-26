import core.TOOL
import core.REGEX
import config.config
import config.regex_reference
import functions.image_process


def extract_pod_data(dn: str) -> list[str] | None:
    """search for information from POD based on DN, only apply to the regular DN pod"""
    res = {"TRANSACTION": [], "AMOUNT": 0}
    try:
        core.TOOL.message(f"start collecting infomation of DN: {dn}")
        pod_file = core.TOOL.file_search_in_directory(
            config.config.global_config["POD"]["save_path"], f"{dn}*.pdf")

        if len(pod_file) == 0:
            return None

        images = functions.image_process.pdf_to_images(pod_file[0])

        pdf_content = functions.image_process.read_images(images)

        pb = core.REGEX.re_find_first(
            config.regex_reference.OCR_FXSJ_PB_REGEX, pdf_content)

        regular_dn = core.REGEX.re_find_first(
            config.regex_reference.OCR_DN_REGEX, pdf_content)

        if regular_dn and pb:
            res["DN"] = regular_dn
            res["PB"] = pb
            for recipient_data in core.REGEX.re_find_all(config.regex_reference.OCR_RECIPIENT_REGEX, pdf_content):
                res["TRANSACTION"].append(
                    (recipient_data[2], recipient_data[4], recipient_data[6]))
                res["AMOUNT"] += int(recipient_data[6])
            core.TOOL.message("information collection complete")
            return res
        core.TOOL.message("information collection failed")
        return None
    except:
        return None
