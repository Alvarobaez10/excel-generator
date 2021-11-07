FROM node:15.11.0-alpine

# set working directory
WORKDIR /app

# install app dependencies
COPY package*.json ./

RUN npm install

COPY . .

EXPOSE 8080

CMD [ "node", "app.js" ]

