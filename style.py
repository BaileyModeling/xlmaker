class Style(object):
    def __init__(self, name, properties):
        self.name = name
        if isinstance(properties, dict):
            self._properties = properties
        else:
            raise TypeError
        self.format = None

    def __str__(self):
        result = self.name + ': \r\n'
        for key, value in self._properties.items():
            result += f'  {key}: {str(value)}; \r\n'
        return result

    def __add__(self, other):
        if self._properties == other._properties:
            return self
        name = self.name + "_" + other.name
        properties = {**self._properties, **other._properties}
        return Style(name, properties)

    def __radd__(self, other):
        if other == 0:
            return self
        else:
            return self.__add__(other)

    def __eq__(self, other):
        return self.__dict__ == other.__dict__

    def __getattr__(self, name):
        return self._properties.get(name)

    def includes(self, other):
        '''True if all properties of other included and identical to self.'''
        result = True
        for key, value in other._properties.items():
            if (
                key not in self._properties or
                value != self._properties[key]
            ):
                result = False
                break
        return result

    def set_property(self, property, value):
        self._properties[property] = value

    def get_property(self, property, default=None):
        return self._properties.get(property, default)

    def get_properties(self):
        return self._properties

    def build(self, workbook):
        if not self.format:
            self.format = workbook.add_format(self.get_properties())

    def get_format(self, workbook):
        if not self.format:
            self.format = workbook.add_format(self.get_properties())
        return self.format

    def append(self, **properties):
        self._properties = {**self._properties, **properties}
    
    def extend(self, name, properties):
        return Style(name, {**self._properties, **properties})

    def copy(self):
        return Style(self.name, self._properties.copy())
