from pyvbaharness.vbasig import parse_declaration, parse_parameter


class TestParameterParsing:
    def test_typed_byval(self):
        p = parse_parameter("ByVal a As Long")
        assert (p.name, p.type_name, p.by_val, p.optional) == (
            "a", "Long", True, False)

    def test_optional_with_default(self):
        p = parse_parameter('Optional ByVal msg As String = "hi"')
        assert p.optional and p.by_val
        assert p.type_name == "String"

    def test_untyped_defaults_to_variant(self):
        p = parse_parameter("value")
        assert p.type_name == "Variant"
        assert not p.by_val

    def test_array_parameter(self):
        p = parse_parameter("items() As Double")
        assert p.is_array
        assert p.type_name == "Double"

    def test_param_array(self):
        p = parse_parameter("ParamArray rest() As Variant")
        assert p.param_array

    def test_signature_carries_params(self):
        sig = parse_declaration(
            "Public Function F(ByVal a As Long, b As Object.Range, "
            "Optional c As Boolean = True) As Long")
        assert [p.type_name for p in sig.params] == [
            "Long", "Object.Range", "Boolean"]
        assert sig.required == 2 and sig.optional == 1
