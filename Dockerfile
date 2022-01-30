FROM node:15.11.0-alpine 

# set working directory
WORKDIR /app

# install app dependencies
COPY package*.json ./

RUN npm install

COPY . .

EXPOSE 3002

CMD [ "node", "./bin/www" ]

