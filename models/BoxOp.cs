using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Box.V2;
using Box.V2.Config;
using Box.V2.JWTAuth;

namespace ReportGen.Models
{
    public static class BoxOp
    {
        static readonly BoxClient UserClient = ConnectBox();

        const string userId = "3983597605"; //My UserId
        const string clientId = "i3dk7jmxb1jtg599gdfovq6u37a5rb22";
        const string clientSecret = "q0eeD9Z7Si38IVl6tXS301Q43T0WobrC";
        const string enterpriseId = "88921423";
        const string paraphrase = "401d18a15f7721e099120df6e433f746";
        const string publicId = "lgwnuv8f";
        const string privateKey = "-----BEGIN ENCRYPTED PRIVATE KEY-----\n" +
                            "MIIFDjBABgkqhkiG9w0BBQ0wMzAbBgkqhkiG9w0BBQwwDgQIzg9EsFnWl1UCAggA\n" +
                            "MBQGCCqGSIb3DQMHBAisYfRCviLNCASCBMizSGH/A74fq4SJNm92rrWrF3SN543N\n" +
                            "vF4uk0rH5+0cSoAPpWE8ZfcNQG8H0B6bsUPkZE4jLf777SQrMMzgB9RDCoN7Mkxl\n" +
                            "f+rQcYIUn2/tpBcdBmqZgwuszmQxbwIPFRooQ7Qvpa6lqyTxJUSxrjL9P/nVTC82\n" +
                            "iZf27/t42NERibFnv834x97eF34Bkne619yvaxx3sdAMx7lFE6e0OwVFPVBTDCUt\n" +
                            "pEh/GoLEKmGw7YKsDYAxuUZBCsJERa31LLnEMok7T0xWAF1xAkCUbFMnXJJi25pN\n" +
                            "qoC/QfGFZsK5Gw2QbeziWlmnaljgPqpN45mGZdnRXWA0VRSc4Nq6CQH9R2qLyUEc\n" +
                            "yWOk/k8rqabWvlvRVxPBSBtIyX3vGCy9jrjASjp6qYNvidkqQsIse2ys5K0hfqzK\n" +
                            "q/XZqgLvUnyEKFk4GQbxif28ZNmWFDTDfGQmawYurDV48qnUjuiRaVHo4UgBp1sr\n" +
                            "2fDmwCCUOVEMooeUCw/bQ5M0KtkG4ckQUhIa6fjKOQmqLqbXqj7S58HHKlNz/q4o\n" +
                            "YbeVH+gXIHWj5AN1AAH98NrIg80pSz5+lxBVM3yfwTVrJrwUIylP1WkhmLWCCwjH\n" +
                            "rkKXwq/jQ8AWs6RLbwq6m4LJbaJuR+wfYZ464C0WjWOkSRUHX95PfIFD6yTchOWH\n" +
                            "4xzRBORDxgWwLvwmQWSioI0Zsvetobzdo8bAgqJCla53N110h2GbXgipoMBEqJ/7\n" +
                            "gGKZW17gr5rcsp6Uz2hv8CEQnB5Bva6K2drFhUZMmeGyuHSSypq2i5OxzPQNsICw\n" +
                            "rwVWJ3bLA1y15L6LwAHNS1ivS53Uha2tV0RBxRXuWSk5vfgnerdtKvrOF2R6s0fI\n" +
                            "6df8H+lWFxttFf7iBHh85taZa+HtTz5iyZQFsjGLR8gN5VSNjDKaShW7evY0cpRF\n" +
                            "luh9JDlHX2T+9zXfTBeHjT7SL9xIJ1fNgwUWl9nK0k27pUJlJiR4DomORb0lqp2Q\n" +
                            "cOZricQuclEKE/GF3ElFiJ+W6oOD52l3CxqDeAN1wMzSVOLmL8c4Zvg3y67vLP04\n" +
                            "NapH30y9obY+fA4pSQ5aPcRBLtDnrWA3CBSTJmJhRQn9I12xVS/l2zXc231fJ935\n" +
                            "YWe9o14DDHNmkVbben/ASy8ciryjErSmNg4qpVg9pAuIP5m7FeXP9XvOja1ApXbL\n" +
                            "59xNevOvZRfYsxL5UUw8TpfuqvxQ2GFBLUM4yYa3W23k5dQ0mC4d709xUUAFQLQt\n" +
                            "JakmjBHZFD53aLA8aesX2Vw1ArP5pqrEV2eCOQZD57i+swbFXHbumNYzcs1p44k9\n" +
                            "COACBNbyKd0JWD0ei9MkR+663mFsmQmgU9hQ4rGZUTesFGZmX1LUBLWGZiABSaS4\n" +
                            "FZ5Fry4OcSnLXjMfC6XLxWi7mL2rmh7mjJO1R9TsyhKim1DKAeCbsnU8/cd9cJed\n" +
                            "eFnAn/upywogRbOHJLu4ZjYQYuEoTgq2kIYDWqCZF0/JK2DOZ0S7vc5/VGXl4HDq\n" +
                            "lB12C9+aS8jS6dm/5l0rrin7ie0Pu98WrPZCtLuNRNlnb4faCI5UI1I+5ZNrSX6K\n" +
                            "mfTCXmMSkO0ekZR7UuL6mgNbDyfpqHLpjJpmO9xXJBo6GKvCU0Dwr+GJgL8puW7C\n" +
                            "1wQ=\n" +
                            "-----END ENCRYPTED PRIVATE KEY-----\n";


        private static BoxClient ConnectBox()
        {
            BoxConfig config = new BoxConfig(clientId, clientSecret, enterpriseId, privateKey, paraphrase, publicId);
            BoxJWTAuth boxJWT = new BoxJWTAuth(config);

            var userToken = boxJWT.UserToken(userId);
            var userClient = boxJWT.UserClient(userToken, userId);

            return userClient;
        }


    }
}
