#!/usr/bin/env bash
# Runner script for Gemini Invoice Document Intelligence

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export PYTHONPATH="$SCRIPT_DIR:$PYTHONPATH"
cd "$SCRIPT_DIR"

MODE=${1:-all}

if [ "$MODE" == "backend" ]; then
    echo "Starting FastAPI backend on port 8000..."
    uvicorn backend.app:app --host 0.0.0.0 --port 8000 --reload
elif [ "$MODE" == "frontend" ]; then
    echo "Starting Streamlit frontend on port 8501..."
    streamlit run frontend/app.py --server.port 8501 --server.address 0.0.0.0
elif [ "$MODE" == "test" ]; then
    echo "Running automated pytest suite..."
    pytest tests -v
else
    echo "Starting both FastAPI Backend and Streamlit Frontend..."
    uvicorn backend.app:app --host 0.0.0.0 --port 8000 &
    BACKEND_PID=$!
    echo "FastAPI started (PID: $BACKEND_PID)"
    
    streamlit run frontend/app.py --server.port 8501 --server.address 0.0.0.0 &
    FRONTEND_PID=$!
    echo "Streamlit started (PID: $FRONTEND_PID)"

    trap "kill $BACKEND_PID $FRONTEND_PID" SIGINT SIGTERM
    wait
fi

