/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Threading.Tasks;

    /// <summary>
    /// Represents a service request that can have multiple responses.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract partial class MultiResponseServiceRequest<TResponse>
    {

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response collection.</returns>
        internal async Task<ServiceResponseCollection<TResponse>> ExecuteAsync()
        {
            ServiceResponseCollection<TResponse> serviceResponses = (ServiceResponseCollection<TResponse>) await this.InternalExecuteAsync();

            if (this.ErrorHandlingMode == ServiceErrorHandling.ThrowOnError)
            {
                EwsUtilities.Assert(
                    serviceResponses.Count == 1,
                    "MultiResponseServiceRequest.Execute",
                    "ServiceErrorHandling.ThrowOnError error handling is only valid for singleton request");

                serviceResponses[0].ThrowIfNecessary();
            }

            return serviceResponses;
        }
    }
}
