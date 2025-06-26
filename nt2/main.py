import export.message as message
import export.SFC_lookup as SFC
import export.send_gr as send_gr
import export.scan_label as scan_label
import export.build_gr as build_gr
import export.send_ITN as ITN
import export.create_FOX_SODN as SODN
import export.shipping_scan as shipping_scan
import gc
print("NICE TOOLs 2 booting")


def main():
    message.message("NICE TOOLS 2 booted")
    while True:
        message.focus_console()
        message.message("NT2: Please select a step to continue\n\n"
                        "\t\t\t1.Send email to cut DN (need a GR file)\n"
                        "\t\t\t2.Create a GR file (from existing one or Feedfile)\n"
                        "\t\t\t3.ITN apply for a DN (need SLI)\n"
                        "\t\t\t4.QR scan for carton ID\n"
                        "\t\t\t5.QR Scan for serial number\n"
                        "\t\t\t6.Product tracking by SN\n"
                        "\t\t\t7.Create Fox SO DN\n"
                        "\t\t\t8.Shipping Scan to SFC")

        user_input = input("\t\t\t")

        match user_input:
            case "quit":
                break

            # case "0":
            #     attention_need = build_report.build_report()
            #     if attention_need is not None:
            #         message.alert(
            #             "following target table rows failed to generate report row")
            #         for row in attention_need:
            #             message.message(f"row: {row[0]}\n"
            #                             f"PB: {row[1]}\t\t\t QTY: {row[2]}\t\t\tDN: {row[3]}\n"
            #                             f"DST: {row[4]}\t\t\tStatus: {row[5]}\t\t\tInvoice: {row[6]}\n"
            #                             f"Owner: {row[7]}\n")

            case "1":
                message.message("please select gr file to apply for a DN")
                send_gr.send_gr()

            case "2":
                message.message(
                    "please select a gr file or feedfile to continue")
                build_gr.build_gr()

            case "3":
                message.message(
                    "please prepare the DN and SLI for ITN application")
                message.focus_console()
                dn = input("please enter the dn for this application\n")
                if dn == "":
                    break
                ITN.send_ITN(dn)

            case "4":
                message.message(
                    "please start scanning the QR code of shipping label")
                qr_data = scan_label.scan_carton_id_from_qr()
                if qr_data != None:
                    for carton_id in qr_data:
                        print(carton_id)
                else:
                    message.message("no data scanned")

            case "5":
                message.message(
                    "please start scanning the QR code of shipping label")
                qr_data = scan_label.scan_sn_from_qr()
                if qr_data != None:
                    for sn in qr_data:
                        print(sn)
                else:
                    message.message("no data scanned")

            case "6":
                message.focus_console()
                product_data = SFC.SFC_product_tracking_by_sn(
                    input("please enter serial number\n"))
                if product_data is None:
                    continue
                message.message(
                    "\n\n\n"
                    f"SN: {product_data.get("SN")}\t\t\t\tNext Station: {product_data.get("Next Station")}\n"
                    f"Work Order: {product_data.get("MO_Number")}\t\t\tPart Number: {product_data.get("Model_Name")}\n"
                    f"PB: {product_data.get("PBR")}\t\t\t\t\tCarton ID: {product_data.get("cartoon_id")}\n")

            case "7":
                while True:
                    message.message(
                        "please enter a DN or the row number of FXSJ vender pool")
                    data = input().replace(" ", "")
                    if data == "quit":
                        break
                    # try:
                    if int(data) > 80000000:
                        # NVDN
                        message.message("Common shipping codes:"
                                        "\nSJ03\tZanker"
                                        "\nBNV021\tCACananda"
                                        "\nSJ04\t9 high-tech, Shenzhen, China"
                                        "\nBNV21SZ\tlonghua, shenzhen, China"
                                        "\nSI06\tBagmane Goldstone Bldg.N.Tower, India"
                                        "\nBNV021IL02\tAshoka Path Airport Rd, India"
                                        "\nSJ07\t13/F  Goodman Tsuen Wan Centre, HongKong"
                                        "\nSJ01\t18/F Goodman Texaco Centre, HongKong"
                                        "\nBNV011HK03\t3/F Harbour View 1, Nvidia Singapore, HongKong"
                                        "\nSJ05\tNo. 8, Kee Hu Road 114, Taipei,Taiwan")
                        message.message(
                            "please enter the shipping code (default SJ03 to Zanker press 'Enter')")
                        shipping_code = input()
                        if shipping_code == "":
                            shipping_code == "SJ03"
                        so, dn = SODN.create_FOX_SODN(data, shipping_code)
                        message.message(
                            f"FXSO and FXDN created for NVDN: {data}\n{so}\n{dn}")
                        break
                    else:
                        # row index
                        so, dn = SODN.create_FXSJ_SODN(int(data))
                        message.message(
                            f"FXSO and FXDN created for FXSJ at row: {data}\n{so}\n{dn}")
                        break
                    # except:
                    #     message.alert("invalid data")
                    #     continue

            case "8":
                while True:
                    message.message(
                        "please enter a DN or the row number of FXSJ vender pool")
                    data = input().replace(" ", "")
                    if data == "quit":
                        break
                    # try:
                    if int(data) > 80000000:
                        # NVDN
                        message.message(
                            "please start scanning shipping label QR")
                        shipping_scan.shipping_scan_NVDN(
                            data, scan_label.scan_carton_id_from_qr())
                        break
                    else:
                        message.message(
                            "please start scanning shipping label QR")
                        shipping_scan.shipping_scan_FXSJ(
                            int(data), scan_label.scan_carton_id_from_qr())
                        break
                    # except:
                        # message.alert("invalid input data")
                        # continue

            # case "7":
            #     cartoon_ids = SFC.SFC_OQC_carton_id_list_lookup_by_sn(
            #         input("please enter serial number\n"))
            #     if cartoon_ids is None:
            #         continue
            #     print("\nCarton ID:")

            #     for carton_id in cartoon_ids[:-1]:
            #         print(carton_id)
            #     print()

            # case "8":
            #     message.focus_console()
            #     serial_numbers = SFC.SFC_PACKING_sn_list_lookup_by_sn(
            #         input("please enter serial number\n"))
            #     if cartoon_ids is None:
            #         continue
            #     print("\nSerial Number:")
            #     for serial_number in serial_numbers[:-1]:
            #         print(serial_number)
            #     print()

            # case "9":
            #     attention_need = build_pod.build_pod()
            #     if attention_need is not None:
            #         message.alert("following files failed to generate POD")
            #         for file in attention_need:
            #             print(file)

            case _:
                continue
        gc.collect()


if __name__ == "__main__":
    main()
