using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointApp.YoutubeModels
{
    public static class Comments
    {

        public class Rootobject
        {
            public string kind { get; set; }
            public string etag { get; set; }

            public string nextPageToken { get; set; }
            public Pageinfo pageInfo { get; set; }
            public Item[] items { get; set; }
        }

        public class Pageinfo
        {
            public int totalResults { get; set; }
            public int resultsPerPage { get; set; }
        }

        public class Item
        {
            public string kind { get; set; }
            public string etag { get; set; }
            public string id { get; set; }
            public Snippet snippet { get; set; }
            public Replies replies { get; set; }
        }

        public class Snippet
        {
            public string videoId { get; set; }
            public Toplevelcomment topLevelComment { get; set; }
            public bool canReply { get; set; }
            public int totalReplyCount { get; set; }
            public bool isPublic { get; set; }
        }

        public class Toplevelcomment
        {
            public string kind { get; set; }
            public string etag { get; set; }
            public string id { get; set; }
            public Snippet1 snippet { get; set; }
        }

        public class Snippet1
        {
            public string authorDisplayName { get; set; }
            public string authorProfileImageUrl { get; set; }
            public string authorChannelUrl { get; set; }
            public Authorchannelid authorChannelId { get; set; }
            public string videoId { get; set; }
            public string textDisplay { get; set; }
            public string textOriginal { get; set; }
            public bool canRate { get; set; }
            public string viewerRating { get; set; }
            public int likeCount { get; set; }
            public DateTime publishedAt { get; set; }
            public DateTime updatedAt { get; set; }
        }

        public class Authorchannelid
        {
            public string value { get; set; }
        }

        public class Replies
        {
            public Comment[] comments { get; set; }
        }

        public class Comment
        {
            public string kind { get; set; }
            public string etag { get; set; }
            public string id { get; set; }
            public Snippet2 snippet { get; set; }
        }

        public class Snippet2
        {
            public string authorDisplayName { get; set; }
            public string authorProfileImageUrl { get; set; }
            public string authorChannelUrl { get; set; }
            public Authorchannelid1 authorChannelId { get; set; }
            public string videoId { get; set; }
            public string textDisplay { get; set; }
            public string textOriginal { get; set; }
            public string parentId { get; set; }
            public bool canRate { get; set; }
            public string viewerRating { get; set; }
            public int likeCount { get; set; }
            public DateTime publishedAt { get; set; }
            public DateTime updatedAt { get; set; }
        }

        public class Authorchannelid1
        {
            public string value { get; set; }
        }

    }
}
