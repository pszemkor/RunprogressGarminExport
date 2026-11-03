#!/usr/bin/env bash
set -euo pipefail

# Configuration
PROJECT_ID="${PROJECT_ID:-garminsync-491522}"
REGION="${REGION:-europe-central2}"
FUNCTION_NAME="${FUNCTION_NAME:-garmin-sync}"
SCHEDULER_JOB="${SCHEDULER_JOB:-garmin-sync-daily-2300}"
SPREADSHEET_ID="${SPREADSHEET_ID:-1r20OoEoFyaAEcm3K-1GfzGCLxc6zk3FqGg823S7xYlA}"
SECRET_NAME="${SECRET_NAME:-garmin-personal-api-key}"
TIMEZONE="Europe/Warsaw"
SCHEDULE="0 23 * * *"

echo "============================================================"
echo "🚀 Deploying GarminSync to Google Cloud Platform"
echo "Project:    ${PROJECT_ID}"
echo "Region:     ${REGION}"
echo "Function:   ${FUNCTION_NAME}"
echo "Schedule:   ${SCHEDULE} (${TIMEZONE})"
echo "============================================================"

# Ensure gcloud uses the right project
gcloud config set project "${PROJECT_ID}"

# 1. Enable required APIs
echo "📦 Enabling required GCP APIs..."
gcloud services enable \
    cloudfunctions.googleapis.com \
    cloudbuild.googleapis.com \
    artifactregistry.googleapis.com \
    run.googleapis.com \
    cloudscheduler.googleapis.com \
    secretmanager.googleapis.com \
    sheets.googleapis.com \
    drive.googleapis.com

# 2. Check / Create Secret for Garmin Personal API Key
if ! gcloud secrets describe "${SECRET_NAME}" >/dev/null 2>&1; then
    echo "🔐 Creating secret '${SECRET_NAME}' in Secret Manager..."
    if [[ -z "${GARMIN_PERSONAL_API_KEY:-}" ]]; then
        read -rsp "Enter your Garmin Personal API Key: " GARMIN_KEY_INPUT
        echo
    else
        GARMIN_KEY_INPUT="${GARMIN_PERSONAL_API_KEY}"
    fi
    echo -n "${GARMIN_KEY_INPUT}" | gcloud secrets create "${SECRET_NAME}" --data-file=-
else
    echo "✅ Secret '${SECRET_NAME}' already exists."
fi

# 3. Create dedicated runtime service account
SA_NAME="garmin-sync-runner"
SA_EMAIL="${SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"

if ! gcloud iam service-accounts describe "${SA_EMAIL}" >/dev/null 2>&1; then
    echo "👤 Creating service account: ${SA_EMAIL}..."
    gcloud iam service-accounts create "${SA_NAME}" \
        --display-name="Garmin Sync Cloud Function Runner"
else
    echo "✅ Service account '${SA_EMAIL}' already exists."
fi

# Grant Secret Manager Secret Accessor to the service account
gcloud secrets add-iam-policy-binding "${SECRET_NAME}" \
    --member="serviceAccount:${SA_EMAIL}" \
    --role="roles/secretmanager.secretAccessor" >/dev/null

echo "⚠️  REMINDER: If your Google Sheet is private, share it as 'Editor' with:"
echo "   👉 ${SA_EMAIL}"

# 4. Deploy Cloud Function (gen2)
echo "⚡ Deploying Cloud Function (${FUNCTION_NAME})..."
gcloud functions deploy "${FUNCTION_NAME}" \
    --gen2 \
    --runtime=python312 \
    --region="${REGION}" \
    --source=. \
    --entry-point=sync_garmin_http \
    --trigger-http \
    --service-account="${SA_EMAIL}" \
    --no-allow-unauthenticated \
    --timeout=540s \
    --memory=512Mi \
    --set-env-vars="USE_PERSONAL_API=true,SPREADSHEET_ID=${SPREADSHEET_ID}" \
    --set-secrets="GARMIN_PERSONAL_API_KEY=${SECRET_NAME}:latest"

# Get Function URL
FUNCTION_URL=$(gcloud functions describe "${FUNCTION_NAME}" --region="${REGION}" --gen2 --format='value(serviceConfig.uri)')
echo "🌐 Cloud Function URL: ${FUNCTION_URL}"

# 5. Setup Cloud Scheduler (23:00 Europe/Warsaw every day)
echo "⏰ Configuring Cloud Scheduler job (${SCHEDULER_JOB})..."
# Grant Invoker permission on Cloud Run underlying service
gcloud run services add-iam-policy-binding "${FUNCTION_NAME}" \
    --region="${REGION}" \
    --member="serviceAccount:${SA_EMAIL}" \
    --role="roles/run.invoker" >/dev/null

if gcloud scheduler jobs describe "${SCHEDULER_JOB}" --location="${REGION}" >/dev/null 2>&1; then
    echo "Updating existing scheduler job..."
    gcloud scheduler jobs update http "${SCHEDULER_JOB}" \
        --location="${REGION}" \
        --schedule="${SCHEDULE}" \
        --time-zone="${TIMEZONE}" \
        --uri="${FUNCTION_URL}" \
        --http-method=POST \
        --oidc-service-account-email="${SA_EMAIL}"
else
    echo "Creating new scheduler job..."
    gcloud scheduler jobs create http "${SCHEDULER_JOB}" \
        --location="${REGION}" \
        --schedule="${SCHEDULE}" \
        --time-zone="${TIMEZONE}" \
        --uri="${FUNCTION_URL}" \
        --http-method=POST \
        --oidc-service-account-email="${SA_EMAIL}"
fi

echo "============================================================"
echo "🎉 Deployment successfully completed!"
echo "Function will execute every night at 23:00 Polish time."
echo "To test manually, run:"
echo "gcloud scheduler jobs run ${SCHEDULER_JOB} --location=${REGION}"
echo "============================================================"
