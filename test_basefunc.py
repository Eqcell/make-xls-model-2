from basefunc import col_to_num, to_xl_ref, to_rowcol
from basefunc import is_equal

def test_is_equal():
    # not tested
    pass
 
def test_basefunc():
    # Excel references
    assert col_to_num("A") == 1
    assert col_to_num("B") == 2
    assert to_xl_ref(1, 1) == "A1"
    assert to_xl_ref(1, 1, base=1) == "A1"
    assert to_xl_ref(0, 0, base=0) == "A1"
    assert to_rowcol("A1") == (1, 1)
    assert to_rowcol("A1") == to_rowcol("a1")
    assert to_rowcol("A1", base=0) == (0, 0)
    assert to_rowcol("AA1") == (1, 27)
