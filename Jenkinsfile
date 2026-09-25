pipeline {
    agent any

    stages {

        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Build Docker Image') {
            steps {
                sh 'docker build -t hackton-app:latest .'
            }
        }

        stage('Stop Old Container') {
            steps {
                sh 'docker stop hackton-app || true'
                sh 'docker rm hackton-app || true'
            }
        }

        stage('Run New Container') {
            steps {
                sh 'docker run -d --name hackton-app -p 8501:8501 hackton-app:latest'
            }
        }
    }
}