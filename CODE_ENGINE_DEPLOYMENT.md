# Deploy to IBM Code Engine

This guide explains how to deploy the Watson AI proxy server to IBM Code Engine.

## Prerequisites

1. IBM Cloud account (sign up at https://cloud.ibm.com)
2. IBM Cloud CLI installed (https://cloud.ibm.com/docs/cli)
3. Docker installed locally (https://www.docker.com/get-started)
4. Code Engine CLI plugin installed

## Installation Steps

### 1. Install IBM Cloud CLI and Code Engine Plugin

```bash
# Install IBM Cloud CLI (if not already installed)
# macOS
curl -fsSL https://clis.cloud.ibm.com/install/osx | sh

# Login to IBM Cloud
ibmcloud login

# Install Code Engine plugin
ibmcloud plugin install code-engine
```

### 2. Create a Code Engine Project

```bash
# Set your region (choose one: us-south, eu-de, jp-tok, etc.)
ibmcloud target -r us-south

# Create a new Code Engine project
ibmcloud ce project create --name watson-proxy-server

# Select the project
ibmcloud ce project select --name watson-proxy-server
```

### 3. Build and Deploy Using Container Registry

#### Option A: Using IBM Cloud Container Registry

```bash
# Create a namespace in Container Registry
ibmcloud cr namespace-add watson-proxy

# Login to Container Registry
ibmcloud cr login

# Build and push the Docker image
docker build -t us.icr.io/watson-proxy/proxy-server:v1 .
docker push us.icr.io/watson-proxy/proxy-server:v1

# Deploy to Code Engine
ibmcloud ce application create \
  --name watson-proxy \
  --image us.icr.io/watson-proxy/proxy-server:v1 \
  --registry-secret icr-secret \
  --port 8080 \
  --min-scale 1 \
  --max-scale 3 \
  --cpu 0.25 \
  --memory 0.5G \
  --env PROJECT_ID=YOUR_PROJECT_ID_HERE
```

#### Option B: Using Docker Hub (Easier)

```bash
# Login to Docker Hub
docker login

# Build and push the image
docker build -t YOUR_DOCKERHUB_USERNAME/watson-proxy:v1 .
docker push YOUR_DOCKERHUB_USERNAME/watson-proxy:v1

# Deploy to Code Engine (public image, no registry secret needed)
ibmcloud ce application create \
  --name watson-proxy \
  --image docker.io/YOUR_DOCKERHUB_USERNAME/watson-proxy:v1 \
  --port 8080 \
  --min-scale 1 \
  --max-scale 3 \
  --cpu 0.25 \
  --memory 0.5G \
  --env PROJECT_ID=YOUR_PROJECT_ID_HERE
```

### 4. Get the Application URL

```bash
# Get the URL of your deployed application
ibmcloud ce application get --name watson-proxy

# Look for the "URL" field in the output
# It will be something like: https://watson-proxy.xxx.us-south.codeengine.appdomain.cloud
```

### 5. Update Your Figma Plugin

Update your Figma plugin's `src/ui.ts` to use the new Code Engine URL instead of `http://localhost:3000`:

```typescript
// Replace all instances of 'http://localhost:3000' with your Code Engine URL
const PROXY_URL = 'https://watson-proxy.xxx.us-south.codeengine.appdomain.cloud';
```

### 6. Test the Deployment

```bash
# Test the health endpoint
curl https://YOUR_CODE_ENGINE_URL.codeengine.appdomain.cloud/

# Should return: {"status":"ok","message":"Watson AI Proxy Server is running"}
```

## Environment Variables

You can set environment variables for your Code Engine application:

```bash
# Set the Watson Project ID
ibmcloud ce application update --name watson-proxy \
  --env PROJECT_ID=your-actual-project-id

# Set the model ID (optional)
ibmcloud ce application update --name watson-proxy \
  --env MODEL_ID=ibm/granite-3-3-8b-instruct
```

## Updating Your Application

```bash
# Build new version
docker build -t YOUR_DOCKERHUB_USERNAME/watson-proxy:v2 .
docker push YOUR_DOCKERHUB_USERNAME/watson-proxy:v2

# Update the application
ibmcloud ce application update \
  --name watson-proxy \
  --image docker.io/YOUR_DOCKERHUB_USERNAME/watson-proxy:v2
```

## Monitoring and Logs

```bash
# View application logs
ibmcloud ce application logs --name watson-proxy --follow

# Get application details
ibmcloud ce application get --name watson-proxy

# List all applications
ibmcloud ce application list
```

## Scaling Configuration

```bash
# Update scaling settings
ibmcloud ce application update --name watson-proxy \
  --min-scale 1 \
  --max-scale 5 \
  --cpu 0.5 \
  --memory 1G
```

## Cost Optimization

- **Min Scale 0**: Set `--min-scale 0` to scale to zero when not in use (saves costs)
- **Right-size resources**: Start with small CPU/memory and increase if needed
- **Monitor usage**: Use IBM Cloud monitoring to track usage and costs

## Troubleshooting

### Check application status
```bash
ibmcloud ce application get --name watson-proxy
```

### View recent logs
```bash
ibmcloud ce application logs --name watson-proxy --tail 100
```

### Test locally first
```bash
# Build the Docker image locally
docker build -t watson-proxy-test .

# Run locally to test
docker run -p 8080:8080 \
  -e PROJECT_ID=your-project-id \
  watson-proxy-test

# Test in browser: http://localhost:8080
```

## Security Best Practices

1. **Don't hardcode secrets**: Use environment variables for API keys and project IDs
2. **Enable HTTPS**: Code Engine provides HTTPS by default
3. **Restrict CORS**: Update CORS settings in `proxy-server.js` to only allow your Figma plugin domain
4. **Use IAM**: Secure your Code Engine application with IAM policies

## Additional Resources

- [IBM Code Engine Documentation](https://cloud.ibm.com/docs/codeengine)
- [Code Engine Pricing](https://www.ibm.com/cloud/code-engine/pricing)
- [Container Registry Documentation](https://cloud.ibm.com/docs/Registry)


