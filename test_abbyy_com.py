"""
Test the minimal ABBYY FineReader 15 OCR sequence via COM.
Run:  python test_abbyy_com.py <input.pdf>
"""
import sys, os, time
import pythoncom
import win32com.client
import win32com.client.gencache

if len(sys.argv) < 2:
    print("Usage: python test_abbyy_com.py <input.pdf> [output_dir]")
    sys.exit(1)

PDF    = sys.argv[1]
if len(sys.argv) >= 3:
    OUTDIR = sys.argv[2]
else:
    OUTDIR = os.path.join(os.path.dirname(os.path.abspath(PDF)), "abbyy_ocr_out")
print(f"Input : {PDF}")
print(f"Output: {OUTDIR}")
os.makedirs(OUTDIR, exist_ok=True)

def main():
    pythoncom.CoInitialize()
    try:
        # EnsureDispatch uses the cached type library (built on first run).
        # If ABBYY is still busy from a prior call, retry on RPC_E_CALL_REJECTED.
        RPC_E_CALL_REJECTED = -2147418111
        for attempt in range(10):
            try:
                fr = win32com.client.gencache.EnsureDispatch(
                    "ABBYY.FineReader15.OCR.Application")
                break
            except Exception as e:
                hr = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                if hr == RPC_E_CALL_REJECTED and attempt < 9:
                    print(f"   ABBYY busy, retrying in 5 s (attempt {attempt+1}) ...")
                    time.sleep(5)
                else:
                    raise
        print(f"COM object OK\n")

        print("1. CreateNewBatchWithOptions ...")
        fr.CreateNewBatchWithOptions(OUTDIR)
        print("   OK")

        print("2. AddImages ...")
        v = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [PDF])
        fr.AddImages(v)
        print("   OK")

        print("3. StartExecute ...")
        fr.StartExecute()
        print("   OK — polling for output (up to 3 min) ...")

        deadline = time.time() + 180
        while time.time() < deadline:
            files = [f for f in os.listdir(OUTDIR) if not f.endswith('.tmp')]
            if files:
                print(f"\n   Output appeared: {files}")
                for f in files:
                    path = os.path.join(OUTDIR, f)
                    print(f"   {f}  ({os.path.getsize(path):,} bytes)")
                break
            sys.stdout.write('.')
            sys.stdout.flush()
            time.sleep(3)
        else:
            print(f"\n   TIMEOUT — dir contents: {os.listdir(OUTDIR)}")

    except Exception as e:
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()
        print("\nDone.")

if __name__ == "__main__":
    main()
