You need to install edge-js in your node-modules directory UNDER nodered (not NPM)

And add this to your settings.js

  functionGlobalContext: {
        edge: require('edge-js')
    },
