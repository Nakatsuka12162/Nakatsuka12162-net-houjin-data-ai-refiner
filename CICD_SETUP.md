# CI/CD Pipeline Setup Guide

## GitHub Secrets Configuration

Go to your GitHub repository → Settings → Secrets and variables → Actions → New repository secret

Add the following secrets:

### Docker Hub
- `DOCKER_USERNAME` - Your Docker Hub username
- `DOCKER_PASSWORD` - Your Docker Hub password or access token

### AWS Credentials
- `AWS_ACCESS_KEY_ID` - AWS IAM user access key
- `AWS_SECRET_ACCESS_KEY` - AWS IAM user secret key

### EC2 Instance
- `EC2_HOST` - EC2 instance public IP or DNS (e.g., `ec2-xx-xx-xx-xx.compute.amazonaws.com`)
- `EC2_USER` - SSH username (usually `ubuntu` or `ec2-user`)
- `EC2_SSH_KEY` - Private SSH key content (entire .pem file content)

## AWS EC2 Setup

### 1. Launch EC2 Instance
```bash
# Amazon Linux 2 or Ubuntu recommended
# Open port 8000 in security group
```

### 2. Install Docker on EC2
```bash
# For Ubuntu
sudo apt update
sudo apt install -y docker.io
sudo systemctl start docker
sudo systemctl enable docker
sudo usermod -aG docker $USER

# For Amazon Linux 2
sudo yum update -y
sudo yum install -y docker
sudo service docker start
sudo usermod -aG docker ec2-user
```

### 3. Create .env file on EC2
```bash
# SSH into EC2 and create .env file
nano ~/.env

# Add your environment variables:
DATABASE_URL=postgresql://...
OPEN_AI_API_KEY=sk-proj-...
API_KEY=...
SPREADSHEET_ID=...
PROJECT_ID=...
PRIVATE_KEY="-----BEGIN PRIVATE KEY-----..."
CLIENT_EMAIL=...
CLIENT_ID=...
SECRET_KEY=your-production-secret
DEBUG=False
```

### 4. Configure SSH Access
```bash
# On your local machine, generate key if needed
ssh-keygen -t rsa -b 4096 -f ~/.ssh/ec2_deploy_key

# Copy public key to EC2
ssh-copy-id -i ~/.ssh/ec2_deploy_key.pub user@ec2-host

# Copy private key content to GitHub secret EC2_SSH_KEY
cat ~/.ssh/ec2_deploy_key
```

## Deployment Flow

1. **Push to GitHub** → Triggers workflow
2. **Build** → Tests and builds Docker image
3. **Push to Docker Hub** → Uploads image with tags (latest + commit SHA)
4. **Deploy to AWS** → SSH to EC2, pulls image, restarts container

## Manual Deployment

If needed, deploy manually:

```bash
# SSH to EC2
ssh -i your-key.pem user@ec2-host

# Pull and run
docker pull your-username/company-research:latest
docker stop company-research || true
docker rm company-research || true
docker run -d \
  --name company-research \
  -p 8000:8000 \
  --env-file ~/.env \
  --restart unless-stopped \
  your-username/company-research:latest
```

## Verify Deployment

```bash
# Check container status
docker ps

# View logs
docker logs -f company-research

# Test endpoint
curl http://your-ec2-ip:8000/health/
```

## Troubleshooting

- **Build fails**: Check Dockerfile and dependencies
- **Docker Hub push fails**: Verify credentials in GitHub secrets
- **SSH fails**: Check EC2_SSH_KEY format and EC2 security group
- **Container won't start**: Check .env file on EC2 and logs
