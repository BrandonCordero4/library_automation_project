import os
import scripts.first as first
import scripts.format as format_script
from flask import Flask, render_template, request, send_file, Response, stream_with_context
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# Single shared state since only one user/job runs at a time
job_state = {"cancelled": False, "ready": False, "running": False}

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")


@app.route("/stream", methods=["POST"])
def stream():
    if job_state["running"]:
        return "Busy", 409
    
    start_row = int(request.form["start_row"])
    end_row   = int(request.form["end_row"])

    excel_path = os.path.join(UPLOAD_FOLDER, f"job_{uuid.uuid4().hex}.xlsx")

    uploaded_file = request.files["spreadsheet"]
    uploaded_file.save(excel_path)

    internal_start = start_row - 2
    num_books      = end_row - start_row + 1

    job_state["cancelled"] = False
    job_state["ready"]     = False
    job_state["running"]   = True

    def generate():
        try:
            for completed, total in first.main(
                excel_file=excel_path,
                start_row=internal_start,
                num_books=num_books
            ):
                yield f"data: {{\"completed\": {completed}, \"total\": {total}}}\n\n"
                if job_state["cancelled"]:
                    raise ConnectionAbortedError("Client disconnected")
            format_script.main(
                file=excel_path,
                start_row=start_row,
                end_row=end_row
            )

            yield f"data: {{\"done\": true, \"path\": \"{os.path.basename(excel_path)}\"}}\n\n"

        except GeneratorExit:
            # Client closed the tab — clean up silently, no trace
            job_state["cancelled"] = True
            if os.path.exists(excel_path):
                os.remove(excel_path)
            raise

        except Exception as e:
            print(f"[STREAM] Job stopped: {e}")
            if os.path.exists(excel_path):
                os.remove(excel_path)
            yield "data: {\"failed\": true}\n\n"
        finally:
            job_state["running"] = False
    return Response(stream_with_context(generate()), mimetype="text/event-stream")



@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(path):
        return "Not ready", 404
    return send_file(path, as_attachment=True, download_name="here_you_go.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", threaded=False)