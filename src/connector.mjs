import { unzipSync, strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { get } from './http.mjs';
export function parts(bytes){const files=unzipSync(bytes);return Object.fromEntries(Object.entries(files).map(([name,value])=>[name,strFromU8(value)]));}
export function paragraphs(xml){const parser=new XMLParser({ignoreAttributes:false,preserveOrder:true});const output=[];function walk(nodes){for(const node of nodes){for(const [name,value]of Object.entries(node)){if(name==='w:t')output.push(value.map(item=>item['#text']??'').join(''));else if(Array.isArray(value))walk(value);}}}walk(parser.parse(xml));return output;}
export async function inspect(base,options={}){const bytes=Buffer.from(await(await get(base,'/agreement.docx?kind='+(options.kind??'normal'))).arrayBuffer());const files=parts(bytes);if(!files['word/document.xml'])throw new Error('正文部件缺失');return {parts:Object.keys(files),text:paragraphs(files['word/document.xml'])};}
