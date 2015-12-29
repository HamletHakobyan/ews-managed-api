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

using System;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    using System.IO;
    using System.Net;
    using System.Threading.Tasks;

    internal abstract partial class ServiceRequestBase
    {
        /// <summary>
        ///  Gets the task which results IEwsHttpWebRequest object from the specified IEwsHttpWebRequest object with exception handling
        /// </summary>
        /// <param name="request">The specified IEwsHttpWebRequest</param>
        /// <returns>An IEwsHttpWebResponse instance</returns>
        protected async Task<IEwsHttpWebResponse> GetEwsHttpWebResponseAsync(IEwsHttpWebRequest request)
        {
            try
            {
                return await request.GetResponseAsync();
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    this.ProcessWebException(ex);
                }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (IOException e)
            {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
        }

        /// <summary>
        /// Validates request parameters, and emits the request to the server.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns>The task which will resplve to response returned by the server.</returns>
        protected Task<IEwsHttpWebResponse> ValidateAndEmitRequestAsync(out IEwsHttpWebRequest request)
        {
            this.Validate();

            request = this.BuildEwsHttpWebRequest();

            if (this.service.SendClientLatencies)
            {
                string clientStatisticsToAdd = null;

                lock (clientStatisticsCache)
                {
                    if (clientStatisticsCache.Count > 0)
                    {
                        clientStatisticsToAdd = clientStatisticsCache[0];
                        clientStatisticsCache.RemoveAt(0);
                    }
                }

                if (!string.IsNullOrEmpty(clientStatisticsToAdd))
                {
                    if (request.Headers[ClientStatisticsRequestHeader] != null)
                    {
                        request.Headers[ClientStatisticsRequestHeader] =
                            request.Headers[ClientStatisticsRequestHeader]
                            + clientStatisticsToAdd;
                    }
                    else
                    {
                        request.Headers.Add(
                            ClientStatisticsRequestHeader,
                            clientStatisticsToAdd);
                    }
                }
            }

            return ValidateAndEmitRequestInternalAsync(request);
        }

        private async Task<IEwsHttpWebResponse> ValidateAndEmitRequestInternalAsync(IEwsHttpWebRequest request)
        {
            DateTime startTime = DateTime.UtcNow;

            IEwsHttpWebResponse response = null;

            try
            {
                response = await this.GetEwsHttpWebResponseAsync(request);
            }
            finally
            {
                if (this.service.SendClientLatencies)
                {
                    int clientSideLatency = (int) (DateTime.UtcNow - startTime).TotalMilliseconds;
                    string requestId = string.Empty;
                    string soapAction = this.GetType().Name.Replace("Request", string.Empty);

                    if (response != null && response.Headers != null)
                    {
                        foreach (string requestIdHeader in ServiceRequestBase.RequestIdResponseHeaders)
                        {
                            string requestIdValue = response.Headers.Get(requestIdHeader);
                            if (!string.IsNullOrEmpty(requestIdValue))
                            {
                                requestId = requestIdValue;
                                break;
                            }
                        }
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append("MessageId=");
                    sb.Append(requestId);
                    sb.Append(",ResponseTime=");
                    sb.Append(clientSideLatency);
                    sb.Append(",SoapAction=");
                    sb.Append(soapAction);
                    sb.Append(";");

                    lock (clientStatisticsCache)
                    {
                        clientStatisticsCache.Add(sb.ToString());
                    }
                }
            }

            return response;
        }
    }
}
