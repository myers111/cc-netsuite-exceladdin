
module.exports = {

    get: async function (params) {

        try {
    
            if (params.options) {
    
                for (var key in params.options) {
    
                    params.path += (params.path.includes('?') ? '&' : '?') + key + '=' + params.options[key];
                }
            }
    
            const options = {
                method: 'GET',
            };
    
            return await doFetch(params.path, options);
        }
        catch (err) {
            
            console.log(err);
    
            return {};
        }
    },
    post: async function(params) {

        try {
    
            const options = {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json'
                },
                body: JSON.stringify(params.options)
            };

            return await doFetch(params.path, options);
        }
        catch (err) {
            
            console.log(err);
    
            return {};
        }
    }
};

async function doFetch(path, options) {

    if (!options) return {};

    try {

        //const response = await fetch('https://cc-netsuite-node.azurewebsites.net/api/quoting-excel/' + path, options);
        const response = await fetch('http://localhost:8080/api/quoting-excel/' + path, options);

        const contentType = response.headers.get("content-type");

        console.log(contentType);

        if (!contentType || !contentType.includes("application/json")) {

            throw new TypeError("Response not JSON!");
        }

        return await response.json(); //extract JSON from the http response
    }
    catch (err) {
        
        console.log(err);

        return {};
    }
}
