# General Attack Classes Reference

This catalog outlines standard vulnerability classes applicable across services, daemons, and libraries.

## 1. Path Traversal & Arbitrary File Access (LFI / Path Traversal)
- **Mechanism**: User-controlled filenames containing `../`, absolute paths `/etc/`, or null bytes bypass directory containment.
- **Sinks**: `open()`, `os.remove()`, `os.path.join()`, `shutil.rmtree()`, `zipfile.ZipFile.extract()`.
- **Verifier Check**: Does the target use `os.path.abspath()` and verify that `commonpath([base_dir, target_path]) == base_dir`?

## 2. Command Injection & Subprocess Vulnerabilities
- **Mechanism**: Untrusted strings concatenated into shell commands.
- **Sinks**: `subprocess.run(shell=True)`, `os.system()`, `os.popen()`.
- **Verifier Check**: Are arguments passed as an explicit list `['tool', arg1, arg2]` with `shell=False`? Can flags be injected (Argument Injection)?

## 3. Server-Side Request Forgery (SSRF)
- **Mechanism**: Server fetches URLs supplied by untrusted users without restricting IP schemes, loopbacks (`127.0.0.1`, `localhost`), or private subnets (`10.0.0.0/8`, `192.168.0.0/16`, `169.254.169.254`).
- **Sinks**: `requests.get()`, `urllib.request.urlopen()`, Chromium headless navigation.

## 4. Insecure Deserialization & Evaluation
- **Mechanism**: Loading serialized objects from untrusted sources.
- **Sinks**: `pickle.loads()`, `yaml.load(Loader=yaml.Loader)` (without SafeLoader), `eval()`, `exec()`.

## 5. Denial of Service (Resource Exhaustion)
- **Mechanism**: Algorithmic complexity attacks (ReDoS), uncompressed archive bombs (Zip Bomb), infinite loops, unbounded in-memory file buffers.
- **Sinks**: `re.compile()`, `zipfile`, unconstrained memory allocations on file upload.
