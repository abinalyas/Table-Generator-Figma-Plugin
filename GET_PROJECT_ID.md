# How to Get Your Watson Project ID

## Method 1: From Watson Studio Dashboard

1. **Go to**: https://dataplatform.cloud.ibm.com/projects?context=wx
2. **Look for your project** in the list (or create a new one)
3. **Click on the project name**
4. **Click "Manage" tab** on the left sidebar
5. **Scroll down** to the "General" section
6. **Copy the "Project ID"** - it looks like: `a1b2c3d4-5678-90ab-cdef-1234567890ab`

## Method 2: Create a New Project (If you don't have one)

1. **Go to**: https://dataplatform.cloud.ibm.com/projects?context=wx
2. **Click "New project" button** (top right)
3. **Select "Create an empty project"**
4. **Enter a project name**: e.g., "Figma Table Generator"
5. **Select a storage service** (or create new Cloud Object Storage)
6. **Click "Create"**
7. **After creation, go to Manage tab** → Copy the **Project ID**

## Method 3: Using IBM Cloud CLI (Advanced)

```bash
# List all projects
ibmcloud resource service-instances --service-name data-science-experience

# Get project details
ibmcloud resource service-instance YOUR_PROJECT_NAME --output json
```

## Quick Test: Check if Project ID is Valid

Once you have the Project ID, test it with curl:

```bash
# Replace with your actual API key and project ID
curl -X POST "https://us-south.ml.cloud.ibm.com/ml/v1/text/chat?version=2023-05-29" \
  -H "Authorization: Bearer $(curl -s -X POST 'https://iam.cloud.ibm.com/identity/token' \
    -H 'Content-Type: application/x-www-form-urlencoded' \
    -d 'grant_type=urn:ibm:params:oauth:grant-type:apikey&apikey=YOUR_API_KEY' | jq -r .access_token)" \
  -H "Content-Type: application/json" \
  -d '{
    "messages": [{"role": "user", "content": "Hello"}],
    "model_id": "ibm/granite-3-3-8b-instruct",
    "project_id": "YOUR_PROJECT_ID"
  }'
```

If you get a 404 error with "container_not_found", the project ID is wrong or not accessible.

## Update proxy-server.js

Once you have YOUR project ID, update line 10 in `proxy-server.js`:

```javascript
const PROJECT_ID = 'YOUR_ACTUAL_PROJECT_ID_HERE';
```

Then restart the proxy server:
```bash
lsof -ti:3000 | xargs kill -9
npm run proxy
```


