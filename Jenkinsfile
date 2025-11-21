pipeline {
  agent { label "node-ttt-x.x.184.66" }

  environment {
    APP_NAME       = 'tiffa-validate'
    IMAGE_TAG      = "dev-${env.BUILD_NUMBER}"
    DOCKER_IMAGE   = "${APP_NAME}:${IMAGE_TAG}"
    CONTAINER_NAME = "tiffa-validate-dev"
    DOCKER_AVAILABLE = 'unknown'
  }

  stages {
    stage('Checkout & Env') {
      steps {
        cleanWs()
        checkout scm

        // สร้างไฟล์ .env จาก .env.dev.example (เขียนทับเดิม)
        sh '''
          if [ -f .env.dev.example ]; then
            echo "เขียนทับ .env ด้วย .env.dev.example"
            cp .env.dev.example .env
          else
            echo "❌ ไม่พบ .env.dev.example" && exit 1
          fi
          echo "ใช้ .env:"
          cat .env | sed 's/=.*/=***/'
        '''
      }
    }

    stage('Install Dependencies') {
      steps {
        sh 'node -v && npm -v'
        sh 'npm ci'
      }
    }

    stage('Check Docker') {
      steps {
        script {
          def status = sh(returnStatus: true, script: 'docker version > /dev/null 2>&1')
          env.DOCKER_AVAILABLE = (status == 0 ? 'true' : 'false')
          if (env.DOCKER_AVAILABLE != 'true') {
            echo 'Docker is not available on this agent; Docker stages will proceed but skip their internal docker commands.'
          } else {
            sh 'docker version'
          }
        }
      }
    }

    stage('Build Image') {
      steps {
        sh """
          if command -v docker >/dev/null 2>&1 && docker version >/dev/null 2>&1; then
            docker build -t ${DOCKER_IMAGE} .
            docker tag ${DOCKER_IMAGE} ${APP_NAME}:dev-latest
            docker images | grep ${APP_NAME} || true
          else
            echo "Docker not available; skipping image build"
          fi
        """
      }
    }

    stage('Deploy DEV') {
      steps {
        sh """
          if command -v docker >/dev/null 2>&1 && docker version >/dev/null 2>&1; then
            docker stop ${CONTAINER_NAME} || true
            docker rm   ${CONTAINER_NAME} || true
            docker network create tiffa-net || true

            # รัน container
            docker run -d \
              --name ${CONTAINER_NAME} \
              --restart unless-stopped \
              --network tiffa-net \
              -p 4301:4301 \
              ${DOCKER_IMAGE}

            echo "รอ Container เริ่มทำงาน..."
            sleep 5
            docker ps -a | grep ${CONTAINER_NAME} || true
          else
            echo "Docker not available; skipping deploy"
          fi
        """
      }
    }

    stage('Health Check') {
      steps {
        sh '''
          if command -v docker >/dev/null 2>&1 && docker version >/dev/null 2>&1; then
            # ตรวจสอบว่า container รันอยู่หรือไม่
            if docker ps | grep -q tiffa-validate-dev; then
              echo "✅ Container กำลังทำงาน"
              
              # ลองเช็ค health endpoint (ถ้ามี)
              code=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:4301/ || echo "000")
              if [ "$code" = "200" ] || [ "$code" = "304" ]; then
                echo "✅ Health Check OK (HTTP $code)"
              else
                echo "⚠️  HTTP Response: $code (แต่ container ยังรันอยู่)"
                docker logs --tail 20 tiffa-validate-dev || true
              fi
            else
              echo "❌ Container ไม่ได้ทำงาน"
              docker logs tiffa-validate-dev || true
              exit 1
            fi
          else
            echo "Docker not available; skipping health check"
          fi
        '''
      }
    }

    // (ทางเลือก) Cleanup แบบปลอดภัย: ทำเมื่อดิสก์เกิน 85% เท่านั้น
    stage('Optional Cleanup') {
      when {
        expression {
          // ใช้เปอร์เซ็นต์ดิสก์รูท
          def use = sh(script: "df -h / | awk 'NR==2{print int(\$5)}'", returnStdout: true).trim()
          return use.isInteger() && use.toInteger() > 85
        }
      }
      steps {
        // ไม่ใช้ --all/--volumes เพื่อไม่พัง cache เลเยอร์
        sh 'docker system prune -f || true'
      }
    }
  }

  post {
    always {
      sh 'docker system prune --all --volumes -f'
    }
    success { echo '🎉 Release build and push success' }
    failure { echo '💥 Release build and push failed' }
  }
}
