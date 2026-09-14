import test from 'node:test';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile} from 'node:fs/promises';

// Exercise the generated artifact: a correct handler alone does not enable
// mobile deployment. Full XML schema validation is a separate release check.
await promisify(execFile)(process.execPath,['scripts/build.mjs']);
const manifest=await readFile('dist/manifest.xml','utf8');
const section=name=>manifest.match(new RegExp(`<${name}>([\\s\\S]*?)</${name}>`))?.[1];

test('generated mobile manifest connects both compose events to the hosted runtime',async()=>{
  const mobile=section('MobileFormFactor');assert.ok(mobile);
  assert.ok(mobile.includes('Type="OnNewMessageCompose" FunctionName="arkOnCompose"'));
  assert.ok(mobile.includes('Type="OnMessageFromChanged" FunctionName="arkOnFromChanged"'));
  assert.ok(mobile.includes('<SourceLocation resid="RuntimeUrl"/>'));
  assert.ok(manifest.includes('id="RuntimeUrl" DefaultValue="https://arklab.github.io/runtime.html"'));
  const runtime=await readFile('dist/runtime.html','utf8');
  assert.ok(runtime.includes('src="runtime.js"'));
  const js=await readFile('dist/runtime.js','utf8');
  assert.ok(js.includes('arkOnCompose'));assert.ok(js.includes('arkOnFromChanged'));
});
test('mobile recovery command has nine icon variants and keeps desktop surfaces and permissions',()=>{
  const mobile=section('MobileFormFactor');
  assert.ok(mobile.includes('xsi:type="MobileMessageReadCommandSurface"'));
  assert.ok(mobile.includes('xsi:type="MobileButton"'));
  assert.ok(mobile.includes('xsi:type="ShowTaskpane"'));
  assert.ok(mobile.includes('<SourceLocation resid="TaskpaneUrl"/>'));
  for (const size of [25,32,48]) for (const scale of [1,2,3]) assert.ok(mobile.includes(`size="${size}" scale="${scale}"`));
  const desktop=section('DesktopFormFactor');
  for (const surface of ['MessageReadCommandSurface','MessageComposeCommandSurface','LaunchEvent']) assert.ok(desktop.includes(`xsi:type="${surface}"`));
  assert.ok(desktop.includes('FunctionName="arkOnCompose"'));assert.ok(desktop.includes('FunctionName="arkOnFromChanged"'));
  assert.ok(manifest.includes('<Id>f0276772-c58a-4917-a7c5-b0d205b1c8d2</Id><Version>1.1.0.0</Version>'));
  assert.ok(manifest.includes('<Permissions>ReadWriteItem</Permissions>'));
  assert.ok(!manifest.includes('ReadWriteMailbox'));
});
