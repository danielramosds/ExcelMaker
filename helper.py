import yaml as yaml_reader


def read_yaml(filename):
    with open(filename, 'r') as stream:
        return yaml_reader.load(stream)


def get_objects_by_name_from_definition(objects, param):
    if "$ref" in param["schema"]:
        object_name = param["schema"]["$ref"].split("/")[len(param["schema"]["$ref"].split("/")) - 1]
    else:
        object_name = param["schema"]["items"]["$ref"].split("/")[len(param["schema"]["items"]["$ref"].split("/")) - 1]

    return get_properties_of_the_object(get_properties_by_property_name(objects, object_name))


def get_base_path(yaml_content):
    return get_properties_by_property_name(yaml_content, "basePath")


def get_all_paths(yaml_content):
    return get_properties_by_property_name(yaml_content, "paths")


def get_objects_definitions(yaml_content):
    return get_properties_by_property_name(yaml_content, "definitions")


def get_params(yaml_content):
    return get_properties_by_property_name(yaml_content, "parameters")


def get_responses(yaml_content):
    return get_properties_by_property_name(yaml_content, "responses")


def get_properties_of_the_object(yaml_content):
    return get_properties_by_property_name(yaml_content, "properties")


def get_properties_by_property_name(yaml_content, property_name):
    return yaml_content.get(property_name)


def is_object(property):
    return isinstance(property, dict) or isinstance(property, set) or isinstance(property, list) or isinstance(property, tuple)
