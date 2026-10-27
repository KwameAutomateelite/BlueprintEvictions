import {revisionServer} from './revision-http.mjs';
import {createDocxPatchClient} from './docx-patch-client.mjs';
const server=revisionServer({secret:process.env.MADDY_REVISION_SECRET,worker:createDocxPatchClient({python:process.env.PYTHON_PATH||'/opt/parser/bin/python'})});
server.requestTimeout=45000;server.headersTimeout=15000;
server.listen(Number(process.env.PORT||8080),'0.0.0.0');
