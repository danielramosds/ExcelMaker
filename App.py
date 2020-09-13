import xlwings as excelWriter
import helper as utils


#file_name = input("Insert file name with extension?")
file_name = "pet.yaml"

yaml_content = utils.read_yaml(file_name)

excelWriter = excelWriter.Book()
first_sheet = excelWriter.sheets[0];
first_sheet.range('A1').value = "Legendas"
first_sheet.range('B1').value = "List"
first_sheet.range('B2').color = (200, 250, 150)

base_path = utils.get_base_path(yaml_content)
paths = utils.get_all_paths(yaml_content)
objectsDefinitions = utils.get_objects_definitions(yaml_content)

for path in paths:
    httpMethod = paths[path]
    for method in httpMethod:
        sheetName = path.replace("/", "_")
        currentSheet = excelWriter.sheets.add(str(method).upper()+" "+sheetName)
        currentSheet.cells(1, 1).value = str(method).upper()+": "+path
        params = utils.get_params(paths[path][method])
        responses = utils.get_responses(paths[path][method])
        currentSheet.range("A3", "H3").value = "INPUT"
        currentSheet.range("A3", "H3").color = (0, 100, 200)
        row = 4
        for param in params:
            inputType = param["in"]
            collumn = 1
            if inputType == "body":
                currentSheet.cells(row, collumn).value = inputType
                objectProperties = utils.get_objects_by_name_from_definition(objectsDefinitions, param)
                for property in objectProperties:
                    if utils.is_object(property):
                        for content in property:
                            if utils.is_object(property):
                                for another_content in content:
                                    if utils.is_object(property):
                                        for just_another in another_content:
                                            if utils.is_object(property):
                                                for just_another2 in just_another:
                                                    if utils.is_object(property):
                                                        for just_another3 in just_another2:
                                                            currentSheet.cells(row,
                                                                               collumn).value = utils.get_properties_by_property_name(
                                                                param, "name")
                                                            collumn = collumn + 1
                                                            currentSheet.cells(row, collumn).value = just_another3

                                                            collumn = collumn + 1
                                                            currentSheet.cells(row,
                                                                               collumn).value = utils.get_properties_by_property_name(
                                                                param, "description")
                                                            row = row + 1
                                                            collumn = collumn - 2
                                                    else:
                                                        currentSheet.cells(row,
                                                                           collumn).value = utils.get_properties_by_property_name(
                                                            param, "name")
                                                        collumn = collumn + 1
                                                        currentSheet.cells(row, collumn).value = just_another2

                                                        collumn = collumn + 1
                                                        currentSheet.cells(row,
                                                                           collumn).value = utils.get_properties_by_property_name(
                                                            param, "description")
                                                        row = row + 1
                                                        collumn = collumn - 2
                                            else:
                                                currentSheet.cells(row,
                                                                   collumn).value = utils.get_properties_by_property_name(
                                                    param, "name")

                                                collumn = collumn + 1
                                                currentSheet.cells(row, collumn).value = just_another

                                                collumn = collumn + 1
                                                currentSheet.cells(row,
                                                                   collumn).value = utils.get_properties_by_property_name(
                                                    param, "description")
                                                row = row + 1
                                                collumn = collumn - 2
                                    else:
                                        currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(
                                            param, "name")

                                        collumn = collumn + 1
                                        currentSheet.cells(row, collumn).value = another_content

                                        collumn = collumn + 1
                                        currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(
                                            param, "description")
                                        row = row + 1
                                        collumn = collumn - 2
                            else:
                                currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param,
                                                                                                               "name")
                                collumn = collumn + 1
                                currentSheet.cells(row, collumn).value = content

                                collumn = collumn + 1
                                currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param,
                                                                                                               "description")
                                row = row + 1
                                collumn = collumn - 2
                    else:
                        currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param, "name")
                        collumn = collumn + 1
                        currentSheet.cells(row, collumn).value = property

                        collumn = collumn + 1
                        currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param, "description")
                        row = row + 1
                        collumn = collumn - 2
            else:
                currentSheet.cells(row, collumn).value = inputType
                collumn = collumn + 1
                currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param, "name")

                collumn = collumn + 1
                currentSheet.cells(row, collumn).value = utils.get_properties_by_property_name(param, "description")
                row = row + 1
                collumn = collumn - 2







excelWriter.save(yaml_content["info"]["title"] + " " + "Information Gap")
