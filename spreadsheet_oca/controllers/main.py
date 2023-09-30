import zipfile
import json
import io

from odoo import http
from odoo.http import request, content_disposition, Controller

class SpreadsheetController(Controller):

    @http.route('/spreadsheet/xlsx', type='http', auth="user", methods=["POST"])
    def get_xlsx_file(self, zip_name, files, **kw):
        files = json.loads(files)
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w', compression=zipfile.ZIP_DEFLATED) as doc_zip:
            for f in files:
                doc_zip.writestr(f['path'], f['content'])
        content = stream.getvalue()
        headers = [
            ('Content-Length', len(content)),
            ('Content-Type', 'application/vnd.ms-excel'),
            ('X-Content-Type-Options', 'nosniff'),
            ('Content-Disposition', content_disposition(zip_name))
        ]

        response = request.make_response(content, headers)
        return response
