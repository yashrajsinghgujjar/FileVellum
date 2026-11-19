if tool_type=='remove_bg':
            try:
                from rembg import remove
                prepared = _prepare_bg_input(raw, max_side=1280)
                output = remove(prepared, session=get_bg_session())
                out = Image.open(BytesIO(output)).convert('RGBA')
            except Exception as exc:
                app.logger.exception('Background removal error in image tool')
                return jsonify(error='Background removal needs rembg[cpu] and its model. '+str(exc)[:180]),503