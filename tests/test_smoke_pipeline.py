def test_imports():
    import Logic.headers_stage1 as h
    from Logic.w10_load_and_unmerge import load_and_unmerge
    from Logic.w15_detect_header import detect_header_and_frame
    from Logic.w20_select_columns import select_and_order_columns
    from Logic.w30_add_sum_rows import append_sum_rows
    from Logic.w40_finalize_save import save_processed
    assert isinstance(h.DESIRED_HEADERS, list)
